"""只测试会改变研究结论的时间、复利、版本和评级去重边界。"""

from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from zipfile import ZipFile

import numpy as np
import pandas as pd

import build_stock_panel as b


class PanelTest(unittest.TestCase):
    def setUp(self):
        self.tmp = TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        for name in ["CACHE", "OUT", "QA"]:
            p = Path(self.tmp.name) / name
            p.mkdir()
            patcher = patch.object(b, name, p)
            patcher.start()
            self.addCleanup(patcher.stop)

    def panel(self, dates):
        return pd.DataFrame({"stock_code": pd.Series(["000001"] * len(dates), dtype="string"),
                             "asof": pd.to_datetime(dates), "market_cap": 1000.0})

    def financial_data(self, current_correction=None, old_correction=None):
        periods = ["2021-12-31", "2022-03-31", "2022-12-31", "2023-03-31"]
        rpt = pd.DataFrame({"Stkcd": "000001", "Reptyp": ["4", "1", "4", "1"],
            "Accper": periods, "Annodt": ["2022-03-30", "2022-04-29", "2023-03-30", "2023-04-29"]})
        base = pd.DataFrame({"Stkcd": "000001", "Accper": periods, "Typrep": "A", "IfCorrect": "0", "DeclareDate": None})
        bs = base.assign(A001000000=[100, 110, 140, 150], A003100000=[50, 60, 70, 80], A003000000=[50, 60, 70, 80])
        inc = base.assign(B002000000=[20, 10, 30, 12], B002000101=[8, 4, 12, 6],
            B001100000=[100, 40, 200, 60], B001300000=[16, 8, 22, 10],
            B001000000=[20, 10, 30, 12], B001216000=[2, 1, 6, 2])
        if current_correction:
            bs.loc[3, ["IfCorrect", "DeclareDate"]] = ["1", current_correction]
        if old_correction:
            inc.loc[1, ["IfCorrect", "DeclareDate"]] = ["1", old_correction]
        return {"年、中、季报基本情况文件114653443": rpt, "资产负债表114132694": bs, "利润表114149331": inc}

    def compute_financial(self, data, dates):
        with patch.object(b, "load", side_effect=lambda p, cols: data[p][cols].copy()):
            return b.financial_panel(self.panel(dates))

    def test_correction_is_not_backdated_and_ttm_uses_comparable_periods(self):
        p = self.compute_financial(self.financial_data(current_correction="2023-08-01"), ["2023-06-30", "2023-09-30"])
        self.assertTrue(np.isnan(p.loc[0, "bm"]))
        self.assertTrue(p.loc[0, "bs_correction_blocked"])
        self.assertAlmostEqual(p.loc[0, "pe_ttm"], 1000 / 14)
        self.assertTrue(np.isnan(p.loc[0, "roe_ttm"]))
        self.assertEqual(p.loc[0, "fin_report_period"], pd.Timestamp("2023-03-31"))
        self.assertAlmostEqual(p.loc[1, "bm"], 80 / 1000)
        self.assertAlmostEqual(p.loc[1, "roa"], (12 + 30 - 10) / ((150 + 110) / 2))

    def test_old_component_correction_blocks_ttm_but_not_current_bm(self):
        p = self.compute_financial(self.financial_data(old_correction="2023-10-01"), ["2023-06-30"])
        self.assertAlmostEqual(p.loc[0, "bm"], 0.08)
        self.assertTrue(np.isnan(p.loc[0, "roa"]))
        self.assertTrue(p.loc[0, "roa_correction_blocked"])
        self.assertTrue(np.isnan(p.loc[0, "pe_ttm"]))
        self.assertTrue(np.isnan(p.loc[0, "ep_ttm"]))

    def test_missing_correction_date_does_not_release_latest_value(self):
        data = self.financial_data()
        data["资产负债表114132694"].loc[3, "IfCorrect"] = "1"
        p = self.compute_financial(data, ["2023-06-30"])
        self.assertTrue(np.isnan(p.loc[0, "bm"]))

    def test_annual_roa_and_publication_cutoff(self):
        p = self.compute_financial(self.financial_data(), ["2023-04-01"])
        self.assertEqual(p.loc[0, "fin_report_period"], pd.Timestamp("2022-12-31"))
        self.assertAlmostEqual(p.loc[0, "roa"], 30 / 120)
        self.assertAlmostEqual(p.loc[0, "pe_ttm"], 1000 / 12)

    def test_parent_profit_valuation_and_additional_fundamentals(self):
        row = self.compute_financial(self.financial_data(), ["2023-06-30"]).iloc[0]
        expected = {"parent_income_ttm": 14, "pe_ttm": 1000 / 14, "ep_ttm": .014,
            "pb": 12.5, "roe_ttm": .2, "leverage": 1 - 80 / 150,
            "asset_growth_yoy": 150 / 110 - 1, "revenue_yoy": .5,
            "nonoperating_profit_share_ttm": 8 / 32, "rd_expense_to_revenue_ttm": 7 / 220, "loss_ttm": 0}
        for field, value in expected.items():
            self.assertAlmostEqual(row[field], value, msg=field)

    def test_loss_zero_and_missing_earnings_have_distinct_meanings(self):
        for parent_ytd, expected_ep, loss in [(-8, 0, 0), (-18, -.01, 1), (np.nan, np.nan, np.nan)]:
            with self.subTest(parent_ytd=parent_ytd):
                data = self.financial_data()
                data["利润表114149331"].loc[3, "B002000101"] = parent_ytd
                row = self.compute_financial(data, ["2023-06-30"]).iloc[0]
                self.assertTrue(pd.isna(row.pe_ttm))
                if pd.isna(expected_ep):
                    self.assertTrue(pd.isna(row.ep_ttm) and pd.isna(row.loss_ttm))
                else:
                    self.assertAlmostEqual(row.ep_ttm, expected_ep)
                    self.assertEqual(row.loss_ttm, loss)

    def test_growth_missing_component_not_zero_and_rd_history_not_backfilled(self):
        data = self.financial_data()
        data["利润表114149331"].loc[1, "B001100000"] = 0
        data["利润表114149331"].loc[1, "B001216000"] = np.nan
        row = self.compute_financial(data, ["2023-06-30"]).iloc[0]
        self.assertTrue(pd.isna(row.revenue_yoy))
        self.assertTrue(pd.isna(row.rd_expense_to_revenue_ttm))
        data = self.financial_data()
        for df in data.values():
            for c in ["Accper", "Annodt"]:
                if c in df:
                    df[c] = (pd.to_datetime(df[c]) - pd.DateOffset(years=5)).dt.strftime("%Y-%m-%d")
        row = self.compute_financial(data, ["2018-06-30"]).iloc[0]
        self.assertTrue(pd.isna(row.rd_expense_to_revenue_ttm))

    def test_rating_one_vote_per_broker_and_publication_lag(self):
        cols = ["Stkcd", "ReportID", "Rptdt", "DeclareDate", "InstitutionID", "Brokern", "Stdrank"]
        rows = [
            ["000001", "1", "2009-05-01", "2009-05-02", "A", "甲", "中性"],
            ["000001", "2", "2009-10-01", "2009-10-02", "A", "甲", "买入"],
            ["000001", "2", "2009-10-01", "2009-10-02", "A", "甲", "买入"],
            ["000001", "6", "2009-10-01", "2009-10-02", "A", "甲", "中性"],
            ["000001", "3", "2009-11-01", "2009-11-02", "B", "乙", "卖出"],
            ["000001", "4", "2009-12-31", "2010-01-02", "B", "乙", "买入"],
            ["000001", "5", "2009-12-31", None, "C", "丙", "买入"],
        ]
        data = pd.DataFrame(rows, columns=cols)
        with patch.object(b, "load", side_effect=lambda p, columns: data.copy() if p.endswith("114629474") else data.iloc[:0].copy()):
            p = b.rating_panel(self.panel(["2009-12-31"]))
        self.assertEqual(p.loc[0, "rec_brokers"], 2)
        self.assertEqual(p.loc[0, "rec_consensus"], 2.5)
        self.assertEqual(p.loc[0, "rec_change"], -0.5)

    def test_six_month_compounding_missing_month_and_unit_conversion(self):
        company = pd.DataFrame({"Stkcd": ["000001"], "Listdt": ["1991-04-03"], "Statco": ["A"],
            "Statdt": ["2026-01-01"], "Markettype": ["4"], "FormerCode": pd.Series([None], dtype="string"), "Crcd": [None], "Commnt": [None]})
        months = pd.date_range("2009-07-31", "2010-12-31", freq="ME").strftime("%Y-%m")
        m = pd.DataFrame({"Stkcd": "000001", "Trdmnt": months, "Clsdt": "28", "Mclsprc": "10", "Msmvttl": "1000", "Msmvosd": "500", "Mretwd": "0.1", "Mretnd": "0.09", "Markettype": "4", "Ndaytrd": "20"})
        m.loc[m.Trdmnt.eq("2010-09"), "Mretwd"] = None
        with patch.object(b, "load", side_effect=lambda p, cols: (company if p.startswith("公司") else m)[cols].copy()):
            p = b.market_panel().set_index("asof")
        self.assertEqual(p.loc["2010-06-30", "market_cap"], 1_000_000)
        self.assertAlmostEqual(p.loc["2010-06-30", "ret6m"], 1.1 ** 6 - 1)
        self.assertTrue(np.isnan(p.loc["2010-12-31", "ret6m"]))
        self.assertTrue(np.isnan(p.loc["2010-06-30", "ret6m_fwd"]))
        self.assertAlmostEqual(p.loc["2009-12-31", "ret6m_fwd"], 1.1 ** 6 - 1)

    def test_cash_dividend_not_used_as_share_adjustment(self):
        d = pd.DataFrame({"Stkcd": ["000001"] * 3, "Disttyp": ["SD", "CA", "RO"],
            "Exdistdt": ["2010-05-01"] * 3, "Annodt": ["2010-04-20"] * 3,
            "Paydt": ["2010-05-03"] * 3, "Amount": ["0.2", "0.5", "0.1"], "Roprc": ["0", "0", "5"], "Markettype": ["4"] * 3})
        factors = pd.DataFrame({"Symbol": ["000001"], "TradingDate": ["2010-05-01"], "CumulateBwardFactor": ["9.0"]})
        with patch.object(b, "load", side_effect=lambda p, cols: (d if p.startswith("分配") else factors)[cols].copy()):
            p = b.distribution_panel(self.panel(["2010-06-30"]))
        self.assertAlmostEqual(p.loc[0, "bonus_share_factor_6m"], 1.2)
        self.assertTrue(p.loc[0, "share_adjustment_review"])
        self.assertEqual(p.loc[0, "price_factor_bwd"], 9)

    def test_xlsx_bad_dimension_does_not_truncate_columns(self):
        xml = b'''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A4"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Stkcd</t></is></c><c r="B1" t="inlineStr"><is><t>Stdrank</t></is></c></row><row r="2"/><row r="3"/><row r="4"><c r="A4" t="inlineStr"><is><t>000001</t></is></c><c r="B4" t="inlineStr"><is><t>buy</t></is></c></row></sheetData></worksheet>'''
        data = BytesIO()
        with ZipFile(data, "w") as z:
            z.writestr("xl/worksheets/sheet1.xml", xml)
        result = b.read_xlsx(data.getvalue(), ["Stkcd", "Stdrank"])
        self.assertEqual(result.to_dict("records"), [{"Stkcd": "000001", "Stdrank": "buy"}])

    def test_capital_effective_date_and_status_publication_both_respected(self):
        c = pd.DataFrame({"Stkcd": ["000001", "000001"], "Shrchgdt": ["2010-01-01", "2010-07-01"],
            "Shrtyp": ["10000", "01000"], "Nshrttl": ["100", "200"], "Nshra": ["50", "100"]})
        s = pd.DataFrame({"Symbol": ["000001", "000001"], "DeclareDate": ["2010-01-01", "2010-07-01"],
            "ChangeDate": ["2010-01-01", "2010-06-01"], "ChangedItem": ["上市状态", "上市状态"],
            "ValueBefore": [None, "正常上市"], "ValueAfter": ["正常上市", "终止上市"], "ReasonID": [None, "S1510"]})
        with patch.object(b, "load", side_effect=lambda p, cols: (c if p.startswith("股本") else s)[cols].copy()):
            result = b.capital_and_status(self.panel(["2010-06-30", "2010-12-31"]))
        self.assertEqual(result.loc[0, "total_shares"], 100)
        self.assertEqual(result.loc[0, "last_status_event"], "正常上市")
        self.assertEqual(result.loc[1, "total_shares"], 200)
        self.assertEqual(result.loc[1, "last_status_event"], "终止上市")


if __name__ == "__main__":
    unittest.main(verbosity=2)
