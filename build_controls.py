"""在 Alpha 面板上计算控制变量（半年频率）。

记号：TE_t 为本期 TotalEquity（基金层面期末净资产，含全部份额，元）；
      TE_{t-1} 为同一基金上一个半年末的 TotalEquity（上一行必须正好是前一个半年末）。

- Size     = ln(TE_t)
- Age      = ln(1 + 月数)，月数 = date 与 成立时间_earliest 的自然月差（成立当月为 0）
- Turnover = [(BuyValue + SellValue) / 2] / [(TE_{t-1} + TE_t) / 2]   半年，不年化
- Expense  = TotalOperatingCost / [(TE_{t-1} + TE_t) / 2]             半年，不年化
- Flow     = [TE_t − (1 + r_t) · TE_{t-1}] / TE_t，r_t = (1 + raw_ret)^6 − 1（本半年累计复权收益）

留空规则：
- 无上一个半年末或其 TE 缺失：Turnover / Expense / Flow 留空；
- 1231 行买卖起算日早于 7 月（年报全年值充当下半年）：Turnover 留空；
- raw_ret 为空（建仓月、不满 6 个月）：Flow 留空。
基金家族规模暂不计算。不缩尾、不截断样本期。

输入：基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha.xlsx（不改）。
输出：基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha_Controls.xlsx。
"""

from pathlib import Path

import numpy as np
import pandas as pd

from excel_style import style_sheet
from merge_alpha import PANEL_FORMATS
from merge_fund_manager import TEXT_COLS

ROOT = Path(__file__).resolve().parent
PANEL = ROOT / "基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha.xlsx"
OUT = ROOT / "基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha_Controls.xlsx"
CONTROLS = ["Size", "Age", "Turnover", "Expense", "Flow"]


def prev_half_end(date: pd.Series) -> pd.Series:
    """YYYY0630 → 上年 1231；YYYY1231 → 同年 0630。"""
    year = date.str[:4].astype(int)
    return np.where(date.str.endswith("0630"), (year - 1).astype(str) + "1231", year.astype(str) + "0630")


def month_diff(date: pd.Series, start: pd.Series) -> pd.Series:
    """date（YYYYMMDD）与 start（YYYY-MM-DD）的自然月差。"""
    d, s = pd.to_datetime(date, format="%Y%m%d"), pd.to_datetime(start)
    return (d.dt.year - s.dt.year) * 12 + (d.dt.month - s.dt.month)


def add_controls(panel: pd.DataFrame) -> pd.DataFrame:
    out = panel.sort_values(["fundid", "date"], ignore_index=True)
    te = out["TotalEquity"]

    # 期初净资产：同一基金上一行，且必须是前一个半年末
    prev = out.groupby("fundid")[["date", "TotalEquity"]].shift()
    out["_has_prev"] = prev["date"].eq(prev_half_end(out["date"]))
    te0 = prev["TotalEquity"].where(out["_has_prev"])
    avg_te = (te0 + te) / 2

    out["Size"] = np.log(te)
    out["Age"] = np.log1p(month_diff(out["date"], out["成立时间_earliest"]))

    # 1231 行买卖起算日早于 7 月：分子是全年发生额，不能当下半年
    out["_annual_as_h2"] = out["date"].str.endswith("1231") & out["BuySellStartdate"].str[5:7].lt("07")
    out["Turnover"] = ((out["BuyValue"] + out["SellValue"]) / 2 / avg_te).mask(out["_annual_as_h2"])
    out["Expense"] = out["TotalOperatingCost"] / avg_te

    r = (1 + out["raw_ret"]) ** 6 - 1
    out["Flow"] = (te - (1 + r) * te0) / te
    out["_te0"] = te0
    return out


def report(out: pd.DataFrame) -> None:
    """打印有值数、留空原因、分位数与极端值。"""
    no_te = out["TotalEquity"].isna()
    no_prev = ~no_te & (~out["_has_prev"] | out["_te0"].isna())
    base = ~no_te & ~no_prev
    reasons = {
        "Size": {"TE缺": no_te},
        "Age": {},
        "Turnover": {
            "TE缺": no_te,
            "无上期TE": no_prev,
            "买卖缺": base & (out["BuyValue"].isna() | out["SellValue"].isna()),
            "年报充当H2": base & out["BuyValue"].notna() & out["_annual_as_h2"],
        },
        "Expense": {"TE缺": no_te, "无上期TE": no_prev, "费用缺": base & out["TotalOperatingCost"].isna()},
        "Flow": {"TE缺": no_te, "无上期TE": no_prev, "raw_ret缺": base & out["raw_ret"].isna()},
    }
    q = [0.01, 0.05, 0.25, 0.5, 0.75, 0.95, 0.99]
    for col in CONTROLS:
        why = " ".join(f"{k}={int(v.sum())}" for k, v in reasons[col].items())
        print(f"{col}: 有值={int(out[col].notna().sum())} 留空={int(out[col].isna().sum())}（{why}）")
        print("  分位:", out[col].quantile(q).round(4).to_dict(), "min", round(out[col].min(), 4), "max", round(out[col].max(), 4))

    # 半年换手率超过 1 很常见（中位数约 1.4），极端值阈值取 10
    show = ["fundid", "date", "TotalEquity", "_te0"]
    for col, mask in [
        ("Turnover", out["Turnover"].gt(10) | out["Turnover"].lt(0)),
        ("Expense", out["Expense"].gt(1) | out["Expense"].lt(0)),
        ("Flow", out["Flow"].abs().gt(1)),
    ]:
        print(f"{col} 极端值行数={int(mask.sum())}，前 8 行：")
        if mask.any():
            print(out.loc[mask, [*show, col]].sort_values(col, key=abs, ascending=False).head(8).to_string(index=False))


def build() -> None:
    panel = pd.read_excel(PANEL, dtype={**{c: str for c in TEXT_COLS}, "BuySellStartdate": str})
    out = add_controls(panel)
    if len(out) != len(panel):
        raise SystemExit("行数变化")
    report(out)

    out = out[[*panel.columns, *CONTROLS]]
    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="fundid_date")
        style_sheet(
            writer.sheets["fundid_date"],
            out.columns,
            text_cols=TEXT_COLS,
            number_formats={**PANEL_FORMATS, **{c: "0.000000" for c in CONTROLS}},
            widths={c: 12 for c in CONTROLS},
        )
    print(f"面板行={len(out)}，写出 {OUT.name}")


if __name__ == "__main__":
    build()
