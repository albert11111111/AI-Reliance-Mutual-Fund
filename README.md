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
