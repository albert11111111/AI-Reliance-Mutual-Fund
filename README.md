# 数据集说明

本目录存放基金研究用数据。每次改数据或口径，在文末「修改记录」追加一条，并同步改对应文件说明。

## 基金名单 → fundid × date 面板

| 文件 | 作用 |
| --- | --- |
| `基金名单.xlsx` | 基金名单。2026-10-31 起含 CSMAR 字段，以及 `成立时间_earliest`、`成立时间_latest`、`发生转型代码重复`、`发生过转型`。 |
| `基金名单_fundid_date.xlsx` | 由名单展开的半年面板。工作表 `fundid_date`。 |
| `build_fundid_date.py` | 生成半年面板。改口径后重跑这一份，不要手改结果表。 |
| `excel_style.py` | 按列名设置 Excel 文本/数值格式与列宽，脚本里不用字母选列。 |
| `基金size.xlsx` | CSMAR 资产负债表节选，含 `TotalEquity`。不改这个文件。 |
| `merge_total_equity.py` | 把 `TotalEquity` 接到面板上，写出新文件。 |
| `基金名单_fundid_date_TotalEquity.xlsx` | 半年面板加上 `TotalEquity`。原面板不改。 |
| `机构持有占比.xlsx` | CSMAR 持有人结构节选。不改这个文件。 |
| `merge_holder_ratio.py` | 把机构/个人/从业人员持有比例接到面板上，写出新文件。 |
| `基金名单_fundid_date_TotalEquity_Holder.xlsx` | 含 `TotalEquity` 与持有人比例的半年面板。 |
| `换手率计算.xlsx` | CSMAR 股票买卖金额节选，含 `BuyValue`、`SellValue`。不改这个文件。 |
| `merge_buy_sell_zero.py` | 在 `…_Holder.xlsx` 上接买卖金额；有报告行时缺侧按 0，一侧减负则买卖都空。 |
| `基金名单_fundid_date_TotalEquity_Holder_BuySell.xlsx` | 含买卖金额的半年面板；买卖成对，无单边空。不改 Holder 原表。 |
| `费用.xlsx` | CSMAR 利润表节选，含 `TotalOperatingCost`（费用合计）。不改这个文件。 |
| `merge_fee.py` | 在 `…_BuySell.xlsx` 上接费用合计；`1231` 用年报减半年报。 |
| `基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee.xlsx` | 再含 `TotalOperatingCost` 的半年面板。不改 BuySell 原表。 |
| `Fund_FundManager.xlsx` | CSMAR 基金经理任期表。不改这个文件。 |
| `merge_fund_manager.py` | 在 `…_Fee.xlsx` 上按时点接经理团队特征（宽表不扩行）。 |
| `基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager.xlsx` | 再含 `Team`/`Female`/`Phd`/`midAge`/`largeAge`。不改 Fee 原表。 |
| `无风险利率.xlsx` | CSMAR 定期整存整取三个月利率变更表（`Irfid02`，单位 %）。不改。 |
| `月度净值.xlsx` | CSMAR 基金份额月度净值。不改。 |
| `build_rf_monthly.py` | 年化利率→月度 `rf_monthly`，并按 `year_month` 接到净值。 |
| `无风险利率_月度.xlsx` | 连续日历月无风险利率。 |
| `月度净值_rf.xlsx` | 月度净值左连接 `rf_monthly`。 |
| `merge_ff_factors.py` | 把三/四/五因子（沪深全 A、流通市值加权）接到 `月度净值_rf`。 |
| `月度净值_rf_三因子.xlsx` | 含 `RiskPremium1`/`SMB1`/`HML1`。 |
| `月度净值_rf_四因子.xlsx` | 再含 Carhart `UMD1`。 |
| `月度净值_rf_五因子.xlsx` | 含 `RMW1`/`CMA1`（`Portfolios=1` 的 2×3）。 |
| `CH因子_鸣石_20261103/` | 鸣石 CH3/CH4 月度因子原文件与说明。不改。 |
| `merge_ch_factors.py` | 把 CH3/CH4 接到 `月度净值_rf`，各出一文件。 |
| `月度净值_rf_CH3.xlsx` | 含 `mktrf`/`SMB`/`VMG` 与 `CH_rf_mon`。 |
| `月度净值_rf_CH4.xlsx` | 再含 `PMO`（用 CH4 自身 `SMB`）。 |
| `rolling_alpha.py` | 24 个月滚动回归（statsmodels `RollingOLS`）计算月度 alpha。 |
| `滚动alpha_CAPM.xlsx` / `…_三因子.xlsx` / `…_四因子.xlsx` / `…_五因子.xlsx` / `…_CH3.xlsx` / `…_CH4.xlsx` | 各模型月度 alpha、历史截距与 beta。 |
| `merge_alpha.py` | 把半年原始收益率和 6 个模型的半年平均 alpha（满 6 个月才填）接到 Manager 面板。 |
| `基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha.xlsx` | 再含 `raw_ret` 与 `capm_alpha`/`ff3_alpha`/`car4_alpha`/`ff5_alpha`/`ch3_alpha`/`ch4_alpha`。不改 Manager 原表。 |
| `build_controls.py` | 在 Alpha 面板上计算控制变量 `Size`/`Age`/`Turnover`/`Expense`/`Flow`。 |
| `基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha_Controls.xlsx` | Alpha 面板再加 5 个控制变量。不改 Alpha 原表。 |
| `merge_evaluation.py` | 把 `基金日净值近半年评价文件*` 的全部指标按半年最后一个交易日接到 Controls 面板末尾。 |
| `基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha_Controls_Eval.xlsx` | Controls 面板再加 CSMAR 近半年评价指标（`EvalDate` 起共 44 列）。不改 Controls 原表。 |

### 成立时间_earliest / 成立时间_latest

这两列在 `基金名单.xlsx` 里，紧挨在 `成立日期` 后面。同一行取 `成立时间` 和 `成立日期` 中更早的一天、更晚的一天，格式仍是文本 `YYYY-MM-DD`。两天相同则两列相同。

当前 3,908 行：3,855 行两列成立日相同，53 行不同。不同的情况来自基金转型，两列谁早谁晚并不固定，所以这里只做最早/最晚，不判断哪一列是转型前。代码重复的两行都保留。

| fundid | 成立时间 | 成立日期 | 成立时间_earliest | 成立时间_latest |
| --- | --- | --- | --- | --- |
| 000006 | 2019-03-19 | 2019-03-19 | 2019-03-19 | 2019-03-19 |
| 000646 | 2014-08-18 | 2019-09-04 | 2014-08-18 | 2019-09-04 |
| 163406 | 2021-01-01 | 2010-04-22 | 2010-04-22 | 2021-01-01 |

### 发生转型代码重复 / 发生过转型

这两列在 `基金转型说明` 后面，取值是整数 `1` 或 `0`。

| 列 | 标 1 | 标 0 | 当前为 1 的行数 |
| --- | --- | --- | --- |
| `发生转型代码重复` | 该 `fundid` 在表里出现多于一行 | 该 `fundid` 只出现一行 | 44 行，22 个代码，每个代码两行都标 1 |
| `发生过转型` | `基金转型说明` 去掉空白后仍有内容 | 说明为空 | 242 |

44 行代码重复的记录，`发生过转型` 也都是 1。另有 198 行说明不为空，但 `fundid` 没有重复。这两列没有删行。

### 面板怎么展开

`基金名单.xlsx` 不改。`build_fundid_date.py` 先按 `fundid` 留一行，再展开写入 `基金名单_fundid_date.xlsx`。`fundid` + `date` 是唯一索引。

重复的 `fundid` 留 `成立时间_earliest` 更早的那一行，用它生成 `date`。长表要保存的列里，这 22 个代码只有成立日不同：14 个的 `成立时间_earliest` 不同，9 个的 `成立时间_latest` 不同（`162107` 两列都不同）。`short_name`、`基金类型ifind`、`基金经理`、`基金ID`、`基金管理公司ID`、`基金管理公司名称` 和两个 0/1 标志在两行之间相同。两行起始日相同（8 个代码）时，日期序列一样，留 `成立时间_latest` 更晚的一行。`162107` 留下的是 `2016-08-24` / `2016-08-25`，另一行的 `成立时间_latest` `2022-01-11` 不进入长表。

每个代码从这条 **`成立时间_earliest` 之后**的第一个半年末开始，之后每个 `0630`、`1231` 一行，直到样本截止日。起始日当天如果正好是半年末，这一天不算，从下一个半年末开始。

样本截止日是 **20260630**。`20261231` 还没到，不进表。

`到期日` 写入长表，并同时用来截断：`0` 视为没有到期日。有真实到期日的行，日期停在到期日当天及之前的最后一个半年末。到期日当天若正好是 `0630` 或 `1231`，这一天保留。到期日早于第一个半年末的行不进面板。

对照：

| fundid | 成立时间_earliest | 到期日 | 第一个 date | 最后一个 date |
| --- | --- | --- | --- | --- |
| 161601 | 2002-09-13 | 0 | 20021231 | 20260630 |
| 000006 | 2019-03-19 | 0 | 20190630 | 20260630 |
| 000209 | 2013-07-17 | 0 | 20131231 | 20260630 |
| 002708 | 2016-06-30 | 0 | 20161231 | 20260630 |
| 163406 | 2010-04-22 | 0 | 20100630 | 20260630 |
| 000646 | 2014-08-18 | 0 | 20141231 | 20260630 |

`163406` 名单里还有一行起始日 `2021-01-01`，长表只用更早的 `2010-04-22`。`000646` 两行起始日都是 `2014-08-18`，长表留下较晚的 `成立时间_latest` `2019-09-04`。

未进入面板的 6 行：

| fundid | 原因 |
| --- | --- |
| 016000 | `成立时间_earliest` 为 2022-09-29，2022-12-21 到期，赶不上 20221231 |
| 019927 | `成立时间_earliest` 为 2024-07-11，2024-10-30 到期，赶不上 20241231 |
| 019010、027293、027622、027921 | `成立时间_earliest` 为 2026-06-30，下一个半年末 20261231 尚未到来 |

当前结果：名单 3,908 行、3,886 个 `fundid`，其中 3,880 个进入面板，共 45,929 行。`date` 从 `20021231` 到 `20260630`，共 48 个半年末。当前名单里最早的 `成立时间_earliest` 是 2002-09-13，没有更早的半年末。

### 列和格式

`基金名单_fundid_date.xlsx` / 工作表 `fundid_date`。`date` 在第二列。下面这些列来自名单，并在每个半年末重复。

| 列 | 格式 | 含义 |
| --- | --- | --- |
| `fundid` | 文本，6 位，保留前导 0 | 基金代码，如 `000006` |
| `date` | 文本，`YYYYMMDD` | 半年末，只取 `0630` 或 `1231`，如 `20190630` |
| `short_name` | 文本 | 基金简称 |
| `基金类型ifind` | 文本 | iFinD 基金类型 |
| `基金经理` | 文本 | 名单里的经理任职区间原文，不是按 `date` 拆开的当期经理 |
| `成立时间_earliest` | 文本，`YYYY-MM-DD` | 该行日期序列的起始日 |
| `成立时间_latest` | 文本，`YYYY-MM-DD` | 该行两个成立日中较晚的一天 |
| `到期日` | 文本；无到期为 `0`，否则 `YYYY-MM-DD` | 名单到期日，同时用于截断面板 |
| `基金ID` | 文本 | CSMAR 基金 ID |
| `基金管理公司ID` | 文本 | CSMAR 基金管理公司 ID |
| `基金管理公司名称` | 文本 | 基金管理公司名称 |
| `发生转型代码重复` | 整数 `1` 或 `0` | 名单里该 `fundid` 是否多于一行 |
| `发生过转型` | 整数 `1` 或 `0` | 名单里转型说明是否有内容 |

排序：`fundid`、`date`。这两列唯一，没有重复组合。空值是缺失，不是 0。读取时把代码和日期按文本读。

```python
pd.read_excel(
    "基金名单_fundid_date.xlsx",
    dtype={"fundid": str, "date": str, "基金ID": str, "基金管理公司ID": str},
)
```

重跑：

```bash
/opt/homebrew/Caskroom/miniconda/base/envs/fund-manager/bin/python build_fundid_date.py
```

样本截止日写在脚本顶部的 `SAMPLE_END`。往后扩到新的半年末时，改这个日期，重跑，并更新本节的行数、截止日期和修改记录。

### TotalEquity

`TotalEquity` 来自《公募基金数据库说明书》表五十六资产负债表（`Fund_FIN_Balance`），中文名是所有者权益合计，单位元。`基金size.xlsx` 前两行是字段中文名和单位，从第三行起是数据。`StateTypeCode` 全部是 `2`（期末）。

接到 `基金名单_fundid_date_TotalEquity.xlsx`，不改原来的两张表。一行面板对上一条报告，四个条件同时成立才填数：

| 面板 | 规模表 | 含义 |
| --- | --- | --- |
| `基金ID` | `FundID` | CSMAR 基金 ID |
| `fundid` | `MasterFundCode` | 基金主代码 |
| `date` 为 `0630` | `EndDate` 同一天，且 `ReportTypeID` 为 `5` | 半年度报告 |
| `date` 为 `1231` | `EndDate` 同一天，且 `ReportTypeID` 为 `6` | 年度报告 |

`ReportTypeID`：`5` 是半年度，`6` 是年度。季报不在这张规模表里。截止日期不是 `06-30` 或 `12-31` 的报告不往半年末上靠。同一天若半年报和年报都有，只取与 `date` 对应的那一种；`519714` 在 `2015-06-30` 两种报告的权益相同，长表用半年度那条。

对上 44,507 行。没对上的 1,422 行 `TotalEquity` 留空，其中 1,366 行是该基金在面板里的第一个 `date`。76 个代码整段都没有可接的期末权益，每个代码在面板里只有一行。空值不是 0。

`000006` 在 `20190630` 的 `TotalEquity` 是 `137005458.29`。

```bash
/opt/homebrew/Caskroom/miniconda/base/envs/fund-manager/bin/python merge_total_equity.py
```

### 持有人比例

来源是《公募基金数据库说明书》表四十六持有人结构（`Fund_HolderStructure`），不是基金评价说明书。《基金评价数据库说明书》里没有这张持有人结构表。`机构持有占比.xlsx` 前两行是字段中文名和单位，从第三行起是数据。

接到 `基金名单_fundid_date_TotalEquity_Holder.xlsx`，不改 `基金名单_fundid_date_TotalEquity.xlsx` 和持有人原表。一行面板对上一条份额报告，四个条件同时成立：

| 面板 | 持有人表 | 含义 |
| --- | --- | --- |
| `基金ID` | `FundID` | CSMAR 基金 ID |
| `fundid` | `Symbol` | 基金代码；面板主代码对应名单里保留的那一份额 |
| `date` 为 `0630` | `EndDate` 同一天，且 `ReportTypeID` 为 `5` | 半年度报告 |
| `date` 为 `1231` | `EndDate` 同一天，且 `ReportTypeID` 为 `6` | 年度报告 |

说明书写明：`TotalInstitutionRatio`、`TotalIndividualRatio`、`TotalPractitionersRatio` 是总产品口径，针对分级基金。本份额口径是 `InstitutionRatio`、`IndividualRatio`、`PractitionersRatio`。

写入面板的三列已经按主产品口径选过：

| 写出列 | 含义 | 单位 |
| --- | --- | --- |
| `InstitutionRatio` | 机构持有占总份额比例 | % |
| `IndividualRatio` | 个人持有占总份额比例 | % |
| `PractitionersRatio` | 基金从业人员持有占总份额比例 | % |
| `HolderRatioScope` | 该期用的口径 | `总产品` 或 `本份额` |

选法：该期三个总产品比例里只要有一个非空，就判定为多份额，三项都取总产品；某一项总产品为空则该项留空，不用本份额去填，避免同一期混口径。三个总产品都空，则取本份额三项。`RelationCode` 仅作核对，不单独决定口径；例如 `000006` 在 `20190630` 为 `S0401`、用本份额 `7.24` / `92.76`，在 `20210630` 为 `S0402`、用总产品 `73.42` / `26.58`。

对上 44,484 行：总产品 23,282 行，本份额 21,202 行。没对上的 1,445 行三列比例和口径都留空。空值不是 0。机构、个人、从业人员分别有数 42,916、44,400、43,503 行。

```bash
/opt/homebrew/Caskroom/miniconda/base/envs/fund-manager/bin/python merge_holder_ratio.py
```

### BuyValue / SellValue

来源是《公募基金数据库说明书》表五十一权益投资买入成本及卖出收入（`Fund_Ptf_BSStockTotal`）。`BuyValue` 是买入股票的成本总额，`SellValue` 是卖出股票的收入总额，单位元。`换手率计算.xlsx` 前两行是字段中文名和单位。

由 `merge_buy_sell_zero.py` 在 `基金名单_fundid_date_TotalEquity_Holder.xlsx` 上接出，写出新文件 `基金名单_fundid_date_TotalEquity_Holder_BuySell.xlsx`。不改 Holder 面板和买卖原表。合并键与规模相同：

| 面板 | 买卖表 | 含义 |
| --- | --- | --- |
| `基金ID` | `FundID` | CSMAR 基金 ID |
| `fundid` | `MasterFundCode` | 基金主代码 |
| `date` 为 `0630` | `EndDate` 同一天，且 `ReportTypeID` 为 `5` | 半年度报告 |
| `date` 为 `1231` | `EndDate` 同一天，且 `ReportTypeID` 为 `6` | 年度报告 |

写出列：

| 列 | 含义 |
| --- | --- |
| `BuySellPeriod` | `H1` 上半年，或 `H2` 下半年/年末首期 |
| `BuySellStartdate` | 该期统计起点；相减得到的 `H2` 为当年 `07-01`，仅用年报时为年报起始日 |
| `BuyValue` | 该期买入股票成本总额，元 |
| `SellValue` | 该期卖出股票收入总额，元 |

当前面板：`H1` 22,508 行，`H2` 20,714 行；买卖都有 43,154，买卖都空 2,775；**无买卖单边空**。

换算规则（point-in-time）：

- `0630`：直接用半年报，记为 `H1`；原表缺买或缺卖视为 `0`
- `1231` 且同年有半年报、起算日相同：年报 − 半年报（两侧缺值都按 `0`），记为 `H2`；起点记为当年 `07-01`
- `1231` 且同年无半年报：直接用年报（缺侧按 `0`），记为 `H2`；起点用年报 `Startdate`
- 年报与半年报 `Startdate` 不一致：该期买卖都留空
- 起算日相同但买入或卖出任一侧减出负值：该期买卖都留空（不单边保留）

面板上仍空的 2,775 行，是买卖表根本没有对应报告，或起算不一致 / 减负整期作废；不是键接错。真正缺报告时的空值不是 `0`。

```bash
/opt/homebrew/Caskroom/miniconda/base/envs/fund-manager/bin/python merge_buy_sell_zero.py
```

### TotalOperatingCost（费用合计）

来源是《公募基金数据库说明书》表五十七业绩与收益分配表（`Fund_FIN_Income`）。`TotalOperatingCost` 是费用合计，单位元；只取 `DATA_TYPE_ID=3`（本期）。`费用.xlsx` 前两行是字段中文名和单位。

由 `merge_fee.py` 在 `基金名单_fundid_date_TotalEquity_Holder_BuySell.xlsx` 上接出，写出新文件 `…_BuySell_Fee.xlsx`。不改 BuySell 面板和费用原表。合并键与买卖相同。换算规则与买卖一致（单列更简单）：

- `0630`：直接用半年报；有报告行但原值为空视为 `0`
- `1231` 且同年有半年报、起算日相同：年报 − 半年报（缺侧按 `0`）
- `1231` 且同年无半年报：直接用年报（缺侧按 `0`）
- 起算日不一致，或减出负值：该期费用留空

```bash
/opt/homebrew/Caskroom/miniconda/base/envs/fund-manager/bin/python merge_fee.py
```

### 基金经理团队特征（时点）

来源是《公募基金数据库说明书》表十基金经理文件（`Fund_FundManager`）。由 `merge_fund_manager.py` 在 `…_BuySell_Fee.xlsx` 上左连接写出 `…_Fee_Manager.xlsx`。不改 Fee 面板和经理原表。合并键：`基金ID`=`FundID`、`fundid`=`MasterFundCode`。

在职判定（point-in-time）：`ServiceStartDate ≤ date`，且 `ServiceEndDate` 为空或 `≥ date`（离职日当天仍算在职）。`PositionState`（1 现任 / 2 前任 / 3 代任）只与离职日做交叉验证，不代替区间过滤。

多名经理时保持 `fundid×date` 一行，追加：

| 列 | 含义 |
| --- | --- |
| `Team` | 该半年末在职经理人数；无人在职则空（不是 0） |
| `Female` | 女性经理人数 / `Team` |
| `Phd` | 博士学历（`DegreeID=P0101`）人数 / `Team` |
| `midAge` | 在职成员证券从业年限（`BusinessDuration`）的中位数；不是出生年龄 |
| `largeAge` | 在职成员从业年限的最大值 |

从业年限缺失时，`midAge`/`largeAge` 只在有年限的成员上计算；全员缺失则这两列空，但 `Team`/`Female`/`Phd` 仍可算。名单里的 `基金经理` 文本列保留，那是静态名单口径。

名单空窗补丁（CSMAR 任期缺口、按名单文案强制在职）：`007449` 的 `20201231`/`20210630` 计入董理；`540007` 的 `20110630` 计入廖志峰。这两人在对应任期行上 `BusinessDuration` 本身为空，故补丁后 `midAge`/`largeAge` 仍空。

`BusinessDuration` 是行内快照，不是随日历自动递增的序列；同基金同名多任期未见下降，仅按姓名汇总会出现“下降”，多半是同名不同人。

```bash
/opt/homebrew/Caskroom/miniconda/base/envs/fund-manager/bin/python merge_fund_manager.py
```

### 月度无风险利率

源字段 `Irfid02` 为定期整存整取三个月年化利率（%）。月度利率：

\[
r_{f,t}^{月}=\left(1+\frac{q_t}{100}\right)^{1/12}-1
\]

`无风险利率.xlsx` 只有利率变更日。先扩成连续日历月：该月末及以前最近一次变更的 `q` 即为当月 `q`（中间无变更的月份沿用上一档）。再写出 `无风险利率_月度.xlsx`（`year_month`、`month_end`、`RateChangeDate`、`q_annual_pct`、`rf_monthly`）。

与 `月度净值.xlsx` 对齐：净值 `TradingDate` 常是当月最后一个**交易日**（约 30% 不是日历月末），不能按精确日期对利率表。用 `TradingDate` 的日历年月 `year_month`（`YYYY-MM`）左连接月度利率，得到 `月度净值_rf.xlsx`。当前净值 274,423 行全部匹配到利率。

```bash
/opt/homebrew/Caskroom/miniconda/base/envs/fund-manager/bin/python build_rf_monthly.py
```

### Fama-French / Carhart 月度因子

源表为 CSMAR 三因子、Carhart 四因子、五因子月度指标。`merge_ff_factors.py` 统一口径后接到 `月度净值_rf.xlsx`，**各出一文件**（不改净值与因子源表）：

| 口径 | 取值 |
| --- | --- |
| 市场 | `MarkettypeID=P9714`（沪深 A 股和创业板和科创板，即沪深全 A） |
| 加权 | 流通市值加权（字段名带 `1`：`RiskPremium1`、`SMB1`、`HML1`、`UMD1`、`RMW1`、`CMA1`） |
| 五因子组合划分 | `Portfolios=1`（2×3，与经典三因子一致） |
| 合并键 | 净值 `year_month` = 因子 `TradingMonth`（`YYYY-MM`） |

`RiskPremium1` 已是市场组合相对三月定存的超额收益。后续滚动回归可用复权净值月增长率（注意单位 %）减 `rf_monthly` 作被解释变量。四因子早期部分月份 `SMB1`/`HML1`/`UMD1` 可能为空（动量形成期等），合并时保留空值并在日志汇报。

```bash
/opt/homebrew/Caskroom/miniconda/base/envs/fund-manager/bin/python merge_ff_factors.py
```

### 鸣石 CH3 / CH4 月度因子

源文件在 `CH因子_鸣石_20261103/`（Liu-Stambaugh-Yuan 中国规模与价值因子）。`merge_ch_factors.py` 按 `year_month`（由 `mnthdt` 转出）左连接到 `月度净值_rf.xlsx`，分别写出 `月度净值_rf_CH3.xlsx`、`月度净值_rf_CH4.xlsx`。

| 列 | 含义 |
| --- | --- |
| `CH_rf_mon` | 鸣石月度无风险利率（一年期存款折月，小数制）；与面板 `rf_monthly`（三月定存）不同 |
| `mktrf` | 市场超额收益，已减 CH 无风险利率，回归时不要再减 |
| `SMB` / `VMG` | 规模、E/P 价值因子；CH4 的 `SMB` 与 CH3 不同，勿混用 |
| `PMO` | 仅 CH4，异常换手率情绪因子 |

因子覆盖 2000-01～2026-08；净值在 1998-03～1999-12 的月份会因子留空。数值为小数制。

```bash
/opt/homebrew/Caskroom/miniconda/base/envs/fund-manager/bin/python merge_ch_factors.py
```

### 滚动回归月度 alpha（CAPM / 三因子 / 四因子 / 五因子 / CH3 / CH4）

`rolling_alpha.py` 分别读取 `月度净值_rf_三因子.xlsx`、`…_四因子.xlsx`、`…_五因子.xlsx`、`月度净值_rf_CH3.xlsx`、`…_CH4.xlsx`，写出 `滚动alpha_CAPM.xlsx`、`滚动alpha_三因子.xlsx`、`…_四因子.xlsx`、`…_五因子.xlsx`、`滚动alpha_CH3.xlsx`、`…_CH4.xlsx`（只保留算出 alpha 的行）。CAPM 只用三因子表里的市场因子 `RiskPremium1`。

对基金 \(j\)、评价月 \(t\)：

\[
R_{j,s}=\widehat a_{j,t-1}+\sum_k\widehat\beta_{j,k,t-1}F_{k,s}+\varepsilon_{j,s},\quad s=t-24,\ldots,t-1
\]

\[
\alpha_{j,t}=R_{j,t}-\sum_k\widehat\beta_{j,k,t-1}F_{k,t}
\]

- 被解释变量 \(R\) = `AdjustedNAVGrowth/100 − rf`，rf 与因子口径一致：CSMAR 四个模型用 `rf_monthly`，CH3/CH4 用 `CH_rf_mon`（`mktrf` 已减该 rf）。
- 因子：CAPM `RiskPremium1`；三因子 `RiskPremium1, SMB1, HML1`；四因子再加 `UMD1`；五因子 `RiskPremium1, SMB1, HML1, RMW1, CMA1`（用五因子表自身的 `SMB1`/`HML1`）；CH3 `mktrf, SMB, VMG`；CH4 `mktrf, SMB, VMG, PMO`（用 CH4 自身的 `SMB`）。
- CH 因子从 2000-01 起，CH3/CH4 只用 2000-01 及以后月份；建仓月仍按基金完整净值历史判定（老基金不会把 2000-01 误当建仓月）。
- 成立月起算：样本从面板成立月（`基金名单_fundid_date.xlsx` 的 `成立时间_earliest` 所在月）开始，之前的净值不进入回归，成立月起第一个非空增长率当建仓月剔除。这主要针对转型基金：面板成立日是转型日，而月度净值从转型前的旧基金开始。实际被剔除成立月前净值的有 5 只：`162107`（金鹰中证 500 指数分级，49 个月）、`164205`（天弘深证成份指数 LOF，56）、`166007`（中欧沪深 300 指数增强 LOF，99）、`550015`（信诚添金分级债券，55）、`020018`（1 个月）。其余净值早于成立月的基金，成立月前只有一行空增长率，不受影响。
- 历史回归含截距；当月 alpha **不减**历史截距；beta 每月重估；窗口截止上月，不含当月。实现上 `RollingOLS(window=24)` 第 \(i\) 行用 \(i-23..i\)，参数后移一行即得 \(i-24..i-1\)。
- 建仓期：每只基金首行增长率本身为空，第一个非空增长率（建仓月）剔除，从第二个起计数。
- 缺月：缺月之后第一行的增长率跨越多个月，剔除；只在日历连续的有效段内滚动，窗口不跨缺月。
- 段内满 24 个月后的下一个月起出 alpha。例：`000006` 剔除 2019-04，用 2019-05～2021-04 估计，首个 alpha 在 2021-05。

| 列 | 含义 |
| --- | --- |
| `ExcessRet` | 当月超额收益 \(R_{j,t}\) |
| `alpha` | 当月 alpha（不减截距） |
| `a_hat` | 上月窗口的历史截距，仅作参考，不参与 alpha |
| `beta_<因子>` | 上月窗口估计的因子暴露 |
| `WindowStart` / `WindowEnd` | 估计窗口首末月（\(t-24\)、\(t-1\)） |

结果：CSMAR 四个模型都是 183,297 行、3,125 只基金，月份 2000-05～2026-06；CH3/CH4 都是 183,142 行、3,125 只基金，月份 2002-01～2026-06（少的 155 行是 2000-05～2001-12 的 alpha）。818 只基金无 alpha，多为净值不足 27 个月的新基金（其中 737 只仍存续到 2026-06）；有缺月的 36 只基金中 3 只因此无 alpha。

```bash
/opt/homebrew/Caskroom/miniconda/base/envs/fund-manager/bin/python rolling_alpha.py
```

### 半年原始收益率与半年平均 alpha（并入面板）

`merge_alpha.py` 读 `…_Manager.xlsx`、`月度净值.xlsx` 和 6 个 `滚动alpha_*.xlsx`，写出 `…_Manager_Alpha.xlsx`。

- 半年归属：1～6 月对应 `YYYY0630`，7～12 月对应 `YYYY1231`。6 个月齐全才填值，否则留空。
- `raw_ret`：半年几何平均月收益 \(\big[\prod_{6\text{个月}}(1+\text{AdjustedNAVGrowth}/100)\big]^{1/6}-1\)（小数制，月度口径，未减无风险利率）。样本规则与滚动回归相同（共用 `rolling_alpha.mark_sample`）：面板成立月起算，剔除建仓月和缺月后首行；但不需要 24 个月估计窗口。
- `*_alpha`：例如 `20240630` 的 alpha = 2024 年 1～6 月月度 alpha 的算术平均（月度口径，不年化）。
- 列名：`capm_alpha`（CAPM）、`ff3_alpha`（三因子）、`car4_alpha`（Carhart 四因子）、`ff5_alpha`（五因子）、`ch3_alpha`、`ch4_alpha`。
- 合并键 `fundid = Symbol`。说明书表六中月度净值的 `Symbol` 是**份额代码**而非主代码，因此脚本先用 `CSMAR原始数据/基金份额信息/` 的份额类别信息表（`Fund_EVL_UnitClassInfo`）核对：净值份额 `(Symbol, FundClassID)` 对应的 `FundID` 必须等于面板 `基金ID`，否则报错。核对结果：面板基金在净值里的份额都是 `Symbol = MasterFundCode` 的主代码份额（`550015` 同一份额类别有转型前后两段历史，转型后一段与面板一致）；面板基金的 C 类等其他份额不在月度净值文件中，不存在一对多。

`raw_ret` 结果：有值 41,324 行；留空 4,605 行，其中 6 个月不全 3,205、该半年无可用月份 1,320（基本是成立当半年只有成立月和建仓月）、整只基金无可用月份 80。`501205`、`501208` 的月度净值比面板成立日晚 3～4 个月才开始，各有 1 期面板行没有任何净值。所有有 alpha 的行都有 `raw_ret`。

alpha 结果（6 列相同）：有值 28,098 行；留空 17,831 行，其中 6 个月不全 2,568（均为基金首个出 alpha 的半年）、该半年无 alpha 13,346（24 个月窗口尚未攒满）、整只基金无 alpha 1,917。CH 因子虽从 2000 年开始，但面板首期是 20021231，所以 CH 两列覆盖与其余四列相同。

注意：

- 时点：`20240630` 行的 alpha 是该半年已实现业绩，与同行规模、费用、经理特征同期；做预测时应用下一期 alpha。
- alpha 只用主代码份额（通常 A 类）净值计算，未合并 C 类等其他份额。
- 面板期内净值全称变化（转型或更名）的 22 只基金未切断滚动窗口，窗口可能跨越变化点：`000646`、`001734`、`003853`、`090019`、`160127`、`160212`、`161818`、`162107`、`162717`、`163406`、`166011`、`168002`、`169102`、`169103`、`169107`、`169109`、`501049`、`501054`、`501066`、`970094`、`970101`、`970113`。其中投资范围明显变化的有 `003853`（2018-04 由金鹰添惠纯债转为金鹰信息产业股票）、`090019`（2018-06 保本转普通混合）、`000646`（2019-09 医疗保健量化转量化优选）、`160212`（2013-02 可分离交易转 LOF）；东方红、兴证资管等多为定开转持有期或集合计划转公募，策略基本延续。

```bash
/opt/homebrew/Caskroom/miniconda/base/envs/fund-manager/bin/python merge_alpha.py
```

### 控制变量

`build_controls.py` 读 `…_Manager_Alpha.xlsx`，写出 `…_Manager_Alpha_Controls.xlsx`（原列不变，末尾加 5 列）。记 \(TE_t\) 为本期 `TotalEquity`（基金层面期末净资产，含全部份额，元），\(TE_{t-1}\) 为同一基金上一个半年末的 `TotalEquity`。

| 列 | 公式 | 说明 |
| --- | --- | --- |
| `Size` | \(\ln TE_t\) | 基金规模 |
| `Age` | \(\ln(1+\text{月数})\) | 月数 = `date` 与 `成立时间_earliest` 的自然月差，成立当月为 0 |
| `Turnover` | \(\dfrac{(BuyValue+SellValue)/2}{(TE_{t-1}+TE_t)/2}\) | 半年换手率，不年化 |
| `Expense` | \(\dfrac{TotalOperatingCost}{(TE_{t-1}+TE_t)/2}\) | 半年费用率，不年化 |
| `Flow` | \(\dfrac{TE_t-(1+r_t)TE_{t-1}}{TE_t}\) | \(r_t=(1+\text{raw\_ret})^6-1\)，本半年累计复权收益（主代码份额） |

- 基金家族规模暂不计算。换手率、费用率直接用当期半年数值，不做过去一年平均。
- 上一行必须正好是前一个半年末（0630 对上年 1231，1231 对同年 0630），否则或其 `TotalEquity` 缺失时，`Turnover`/`Expense`/`Flow` 留空。
- 1231 行 `BuySellStartdate` 早于 7 月（年报全年买卖额充当下半年）时 `Turnover` 留空：218 行，其中 211 行在 2010 年（CSMAR 缺这些基金 2010 年半年报买卖额）。费用对应的 717 行「年报充当下半年」都没有上期净资产，已自动留空。
- `raw_ret` 为空时 `Flow` 留空。不缩尾，不截断样本期。

结果：

| 变量 | 有值 | 留空原因 | 中位数 | 1% / 99% 分位 |
| --- | --- | --- | --- | --- |
| `Size` | 44,507 | TE 缺 1,422 | 19.97 | 16.06 / 23.34 |
| `Age` | 45,929 | — | 3.76 | 0 / 5.41 |
| `Turnover` | 39,363 | TE 缺 1,422；无上期 TE 3,804；买卖缺 1,122；年报充当 H2 218 | 1.36 | 0.11 / 8.08 |
| `Expense` | 40,643 | TE 缺 1,422；无上期 TE 3,804；费用缺 60 | 0.0091 | 0.0046 / 0.0283 |
| `Flow` | 40,681 | TE 缺 1,422；无上期 TE 3,804；`raw_ret` 缺 22 | −0.080 | −2.55 / 0.83 |

数据异常：

- `Turnover` 大于 10 的有 188 行，最大 34.9（`009128` 20260630），多为规模较小或规模快速变化的基金。
- `Expense` 为负 3 行，均在 20090630：`166001`、`202003`、`519694`，源表 `TotalOperatingCost` 本身为负。
- `Flow` 分母是期末规模，大额赎回时无下界：小于 −1 的有 1,936 行，最小 −487.9（`003190` 20180630，规模从 5.06 亿降到 101 万）。上界不超过 1。实证前需考虑缩尾。
- `TotalEquity` 为全部份额合计，而 \(r_t\) 只是主代码份额收益，作为近似。
- 时点：这些变量按报告期末归属，未按公告日过滤；用于预测时须滞后一期。

```bash
/opt/homebrew/Caskroom/miniconda/base/envs/fund-manager/bin/python build_controls.py
```

### CSMAR 近半年评价指标（基金日净值近半年评价文件）

`merge_evaluation.py` 读 19 个 `基金日净值近半年评价文件*/FUND_EVL_DAYAHAY.xlsx`（基金评价数据库「基金日净值近半年评价文件」），写出 `…_Controls_Eval.xlsx`（原列不变，末尾加 `EvalDate` 及源表除 `TradingDate`/`Symbol`/`FullName`/`FundClassID` 外的全部列）。源表每行是某份额在 `TradingDate` 这天、回看近半年日净值算出的评价指标：平均收益率、标准差、Sharpe、Beta、詹森 Alpha、Treynor、TM / CL 择时与选股、区间收益率，以及每项的同类排名和参与排名基金数；市场组合按类别（股票型 F0401 上证综指，债券型 F0402 上证国债，混合型 F0403 两者各 50%）。

- 源文件：17 个不同批次，每批覆盖当年 7 月至次年 6 月，合起来为 2009-07 至 2026-06，无缺年。`123209348` 有三份完全相同的副本（MD5 相同），只读一次。
- 代码：面板 `fundid` = 源表 `Symbol`（说明书中为基金代码，即份额代码）。源表 3,656 个 `Symbol` 的 `FundClassID` 与 `月度净值.xlsx` 一致，即都是面板主代码份额。
- 日期：面板 `date`（0630 / 1231）取该半年最后一个交易日的那一行，记在 `EvalDate`。CSMAR 在 6/30、12/31 落在周末时仍给部分基金记一行（如 2018-06-30 周六只有 483 只，2018-06-29 周五有 737 只），这些不是交易日，用了会使排名不在同一截面，故只取工作日中的最大日期。共 10 个半年末因此提前到周五（如 20180630 → 2018-06-29）。
- 同一 `Symbol` × 半年末在源表中唯一，左连接，行数不变。

结果：45,929 行中有值 41,320，留空 4,609：评价文件起点（20091231）之前 922；整只基金不在评价文件 253（基本是 20260630 刚成立的新基金）；早于该基金首次评价 3,434（成立不满约半年，CSMAR 尚未评价）；中间缺失 0。源表 `CategoryID` 只有 S0601 普通股票型 8,638、S0604 混合型 32,679、S0602 债券型 3 行。

数据异常：

- `RangeReturn` 与 `raw_ret` 折算的半年收益 \((1+\text{raw\_ret})^6-1\) 相关系数 0.9955，差值中位数 0.79 个百分点，最大 31 个百分点。差异来自窗口：CSMAR 的「近半年」从窗口内第一个日净值算起。建仓期只披露周净值的新基金窗口更短，如 `025686` 20260630 的区间起点是 2026-01-09 的净值 1.1763，`RangeReturn` 为 118%，而 2025-12-31 至 2026-06-30 实际涨了 149%。
- `001039` 20151231 的 `RangeReturn` 为 +12.3%，而 `AvgReturn` 为负，月度净值 6/30→12/31 也是 −8.4%，CSMAR 表内口径不一致。
- 极端值：`Treynor` 最大 276.1（Beta 接近 0 时分母很小），`Beta` 最小 −0.89、最大 8.73，均未缩尾。
- 时点：指标只用 `EvalDate` 及之前的日净值，半年末即可得到。排名为当日同类截面排名。

```bash
/opt/homebrew/Caskroom/miniconda/base/envs/fund-manager/bin/python merge_evaluation.py
```

## 修改记录

- 2026-10-30：由 `基金名单.xlsx` 生成 `基金名单_fundid_date.xlsx`。`date` 为成立日之后的 `0630`/`1231`，有到期日则截到到期日，样本截止 `20260630`。去掉 25 行完全重复的名单记录。3,937 只基金、46,897 行。该面板对应的是当时的名单，尚未用后来补入 CSMAR 字段的名单重做。
- 2026-10-31：在更新后的 `基金名单.xlsx`（3,908 行，已含 CSMAR 字段）中，于 `成立日期` 后新增 `成立时间_earliest`、`成立时间_latest`，分别取 `成立时间` 与 `成立日期` 的较早、较晚一天。3,855 行两日相同，53 行不同。未去重，未重做半年面板。
- 2026-10-31：在 `基金转型说明` 后新增 `发生转型代码重复`、`发生过转型`。前者对重复 `fundid` 的每一行标 1（44 行、22 个代码），后者对转型说明非空标 1（242 行）。未删行。
- 2026-10-31：用更新后的 `基金名单.xlsx` 重做 `基金名单_fundid_date.xlsx`。日期从该行 `成立时间_earliest` 之后的第一个 `0630`/`1231` 起，`到期日` 为 `0` 视为未到期，样本截止 `20260630`。长表只保留 `fundid`、`date` 和指定的 10 列。重复 `fundid` 的两行各自展开。3,902 行名单进入面板，3,880 个代码，46,214 行；6 行没有可用半年末。`基金名单.xlsx` 未改。
- 2026-10-31：长表改为 `fundid` + `date` 唯一。重复 `fundid` 只留 `成立时间_earliest` 更早的一行来生成日期；起始日相同时留 `成立时间_latest` 更晚的一行。3,880 个代码、45,929 行。`基金名单.xlsx` 未改。
- 2026-11-02：按说明书表五十六，把 `基金size.xlsx` 的期末 `TotalEquity` 接到新文件 `基金名单_fundid_date_TotalEquity.xlsx`。键为 `基金ID`=`FundID`、`fundid`=`MasterFundCode`，`0630` 对半年度报告、`1231` 对年度报告，截止日期必须是同一天。对上 44,507 行，1,422 行留空。原面板和规模表未改。
- 2026-11-02：按说明书表四十六，把 `机构持有占比.xlsx` 接到新文件 `基金名单_fundid_date_TotalEquity_Holder.xlsx`。键为 `基金ID`=`FundID`、`fundid`=`Symbol`，半年末对半年度/年度报告。有总产品比例的期用总产品三项，否则用本份额三项；`HolderRatioScope` 标记口径。对上 44,484 行。原面板和持有人表未改。
- 2026-11-02：按说明书表五十一，把 `换手率计算.xlsx` 的 `BuyValue`、`SellValue` 接到新文件 `基金名单_fundid_date_TotalEquity_Holder_BuySell.xlsx`。合并键与规模相同；另保留 `BuySellStartdate`。对上 43,222 行。年报买卖额多为全年累计，不能直接当作下半年。原面板和买卖表未改。
- 2026-11-02：`1231` 的买卖改为年报减同年半年报；`0630` 仍用半年报。起算日不一致或减出负值留空。面板买入有数 41,273、卖出 41,223。原面板和买卖表未改。
- 2026-11-02：无同年半年报的 `1231` 改为直接用年报原值。起算日相同但减出负值仍留空：买入 3 条（`007110` 2021、`163110` 2019、`400020` 2015），卖出 2 条（`163110` 2019、`019309` 2025 仅差 0.04 元）。原面板和买卖表未改。
- 2026-11-02：删除旧 `merge_buy_sell.py`。现行 `merge_buy_sell_zero.py` 直接在 `…_Holder.xlsx` 上写出 `…_BuySell.xlsx`：有报告行时缺侧按 `0`；起算日相同且任一侧减负则该期买卖都空。买卖都有 43,154、都空 2,775，无单边空。减负整期留空 4 条（`007110` 2021、`163110` 2019、`400020` 2015、`019309` 2025）。Holder 与买卖原表未改。
- 2026-11-02：`build_fundid_date.py` 长表新增 `到期日`（仍用于截断，规则不变）。重跑规模 / 持有人 / 买卖三步，行数与匹配数不变；`基金名单.xlsx` 未改。
- 2026-11-02：写出格式改为 `excel_style.style_sheet` 按列名设置，去掉各脚本里的 Excel 列字母。
- 2026-11-02：按说明书表五十七，把 `费用.xlsx` 的 `TotalOperatingCost` 接到新文件 `…_BuySell_Fee.xlsx`。`1231` 为年报减半年报，规则同买卖。费用有数 44,459、空 1,470；起算不一致 53、减负留空 11。BuySell 与费用原表未改。
- 2026-11-03：按说明书表十，把 `Fund_FundManager.xlsx` 按时点接到新文件 `…_Fee_Manager.xlsx`。写出 `Team`/`Female`/`Phd`/`midAge`/`largeAge`（从业年限中位数/最大值）；`PositionState` 与离职日交叉验证通过。有团队 45,926、无人在职 3；Fee 与经理原表未改。
- 2026-11-03：名单空窗补丁——`007449` 两期补董理、`540007` 一期补廖志峰；无人在职清零。并检查从业年限：同基金同名多任期无下降；仅按姓名有下降，疑同名。
- 2026-11-03：`build_rf_monthly.py` 将整存整取三月期年化利率换成月度 `rf_monthly`（缺月沿用），按 `year_month` 接到净值，写出 `无风险利率_月度.xlsx` 与 `月度净值_rf.xlsx`。
- 2026-11-03：`merge_ff_factors.py` 按 `P9714`+流通市值加权把三/四/五因子接到净值，分别写出 `月度净值_rf_三因子.xlsx`、`…_四因子.xlsx`、`…_五因子.xlsx`。
- 2026-11-03：`merge_ch_factors.py` 把鸣石 CH3/CH4 月度因子接到净值，写出 `月度净值_rf_CH3.xlsx`、`…_CH4.xlsx`。
- 2026-11-03：新增 `rolling_alpha.py`，对三/四/五因子做 24 个月滚动回归（含截距、窗口截止上月），按上月 beta 计算当月 alpha（不减截距），写出 `滚动alpha_三因子.xlsx`、`…_四因子.xlsx`、`…_五因子.xlsx`。剔除建仓月与缺月后首行，窗口不跨缺月。各 183,557 行、3,125 只基金。净值原表未改。
- 2026-11-03：`rolling_alpha.py` 新增 CAPM（只用三因子表的 `RiskPremium1`），写出 `滚动alpha_CAPM.xlsx`，样本规则同上，183,557 行、3,125 只基金。
- 2026-11-03：`rolling_alpha.py` 新增 CH3/CH4（被解释变量减 `CH_rf_mon`，样本从 2000-01 截断），写出 `滚动alpha_CH3.xlsx`、`…_CH4.xlsx`，各 183,402 行、3,125 只基金，首个 alpha 在 2002-01。建仓月改为只按增长率非空判定（不受因子缺失影响），CSMAR 四个模型结果不变。
- 2026-11-03：`rolling_alpha.py` 样本改为从面板成立月起算（转型基金不再用转型前旧基金净值估 beta），5 只基金受影响，重跑 6 个模型：CSMAR 各 183,297 行，CH 各 183,142 行。
- 2026-11-03：新增 `merge_alpha.py`，6 个模型的月度 alpha 按半年取平均（满 6 个月才填），经份额类别信息表核对 `fundid`/`基金ID` 后左连接到 Manager 面板，写出 `…_Manager_Alpha.xlsx`。45,929 行不变，每列有值 28,098。Manager 原表与 alpha 表未改。
- 2026-11-03：`merge_alpha.py` 新增 `raw_ret`（`月度净值.xlsx` 半年 6 个月几何平均月收益，样本规则同滚动回归但无 24 个月窗口），有值 41,324。`rolling_alpha.py` 把样本判定抽成 `mark_sample` 供两处共用，alpha 结果不变。
- 2026-11-03：新增 `build_controls.py`，在 Alpha 面板上计算 `Size`、`Age`、`Turnover`、`Expense`、`Flow`（半年口径，家族规模暂不算），写出 `…_Alpha_Controls.xlsx`；年报充当下半年的 218 行换手率留空。`merge_alpha.py` 的格式字典抽成 `PANEL_FORMATS` 供复用。Alpha 原表未改。
- 2026-11-03：新增 `merge_evaluation.py`，把 17 批（去掉 2 份重复副本）基金日净值近半年评价文件的全部指标，按 `fundid`=`Symbol`（`FundClassID` 与月度净值一致）及半年最后一个交易日（剔除周末记录）接到 Controls 面板末尾，写出 `…_Controls_Eval.xlsx`。有值 41,320，留空 4,609。Controls 原表与评价文件未改。
