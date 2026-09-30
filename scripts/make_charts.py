import os, datetime as dt
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
import pandas as pd
import numpy as np
from data import rows

OUT = os.environ.get("CHART_OUT", "../charts")
os.makedirs(OUT, exist_ok=True)

for f in fm.findSystemFonts():
    if "NotoSansCJK" in f:
        fm.fontManager.addfont(f)
plt.rcParams["font.family"] = "Noto Sans CJK JP"
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.dpi"] = 150
C1, C2, C3, C4, C5 = "#1F4E78", "#C0504D", "#9BBB59", "#F79646", "#8064A2"
GREY = "#7F7F7F"

def save(fig, name):
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, name), bbox_inches="tight")
    plt.close(fig)

df = pd.DataFrame(rows, columns=["img","status","date","layout","area","floor","tf","orient","total","unit","decor"])
df["date"] = pd.to_datetime(df["date"])
df["rooms"] = df.layout.str[0].astype(int)
df["q"] = df.date.dt.to_period("Q").astype(str)
df["half"] = df.date.dt.year.astype(str) + "H" + ((df.date.dt.month > 6) + 1).astype(str)
def bin_(a):
    if a < 60: return "60㎡以下"
    if a < 90: return "75-77㎡"
    if a < 100: return "94-97㎡"
    if a < 120: return "106-117㎡"
    if a < 150: return "124㎡"
    return "150㎡以上"
df["bin"] = df.area.apply(bin_)
df["comp"] = (df.rooms == 3) & df.area.between(105, 117) & (df.date >= "2026-03-01")

quarters = ["2025Q1","2025Q2","2025Q3","2025Q4","2026Q1","2026Q2","2026Q3"]
g = df.groupby("q")
wavg = (g.total.sum() * 10000 / g.area.sum()).reindex(quarters)
cnt = g.size().reindex(quarters)
big = df[(df.rooms == 3) & (df.bin == "106-117㎡")]
bigmed = big.groupby("q").unit.median().reindex(quarters)

# ---- 图1 季度单价走势
fig, ax = plt.subplots(figsize=(8, 4.2))
ax.plot(quarters, wavg.values, marker="o", color=C1, lw=2.2, label="全样本 面积加权均价")
ax.plot(quarters, bigmed.values, marker="s", color=C2, lw=2.2, label="106-117㎡三房 中位数")
for x, y in zip(quarters, wavg.values):
    ax.annotate(f"{y:,.0f}", (x, y), textcoords="offset points", xytext=(0, -14), ha="center", fontsize=8, color=C1)
for x, y in zip(quarters, bigmed.values):
    ax.annotate(f"{y:,.0f}", (x, y), textcoords="offset points", xytext=(0, 8), ha="center", fontsize=8, color=C2)
ax.axvspan(2.5, 3.5, color="#FBE5D6", alpha=0.6)
ax.text(3, 14700, "2025Q4 单季跌约14%\n成交量放大到23套", ha="center", fontsize=8.5, color=C4)
ax.set_ylim(10000, 15200)
ax.set_ylabel("元/㎡")
ax.set_title("图1  国创光谷上城 季度成交单价走势（2025Q1–2026Q3）")
ax.grid(axis="y", alpha=0.3)
ax.legend(loc="lower left", fontsize=9)
save(fig, "01_季度单价走势.png")

# ---- 图2 季度成交量
fig, ax = plt.subplots(figsize=(8, 3.6))
bars = ax.bar(quarters, cnt.values, color=[C1]*3 + [C4] + [C1]*3)
for b, v in zip(bars, cnt.values):
    ax.text(b.get_x() + b.get_width()/2, v + 0.4, str(v), ha="center", fontsize=9)
ax.set_ylabel("成交套数（截图样本）")
ax.set_title("图2  季度成交量：跌价换来的是放量")
ax.grid(axis="y", alpha=0.3)
ax.text(0.99, 0.95, "注：2025年8月、2026年2月无记录，可能有截图遗漏", transform=ax.transAxes, ha="right", va="top", fontsize=8, color=GREY)
save(fig, "02_季度成交量.png")

# ---- 图3 分面积段
halves = ["2025H1", "2025H2", "2026H1", "2026H2"]
bins = ["75-77㎡", "94-97㎡", "106-117㎡"]
fig, ax = plt.subplots(figsize=(8, 4.2))
w = 0.2
x = np.arange(len(bins))
cols = [C1, C5, C3, C2]
for i, h in enumerate(halves):
    vals = [df[(df.bin == b) & (df.half == h)].unit.median() for b in bins]
    ns = [len(df[(df.bin == b) & (df.half == h)]) for b in bins]
    bs = ax.bar(x + (i - 1.5) * w, vals, w, label=h, color=cols[i])
    for b_, v, n in zip(bs, vals, ns):
        ax.text(b_.get_x() + b_.get_width()/2, v + 100, f"{v:,.0f}\n({n}套)", ha="center", fontsize=7)
ax.set_xticks(x); ax.set_xticklabels(bins)
ax.set_ylim(9000, 15500)
ax.set_ylabel("单价中位数（元/㎡）")
ax.set_title("图3  分面积段单价：各户型同步下台阶，2026下半年略有回升")
ax.legend(fontsize=9, ncol=4, loc="upper left")
ax.grid(axis="y", alpha=0.3)
save(fig, "03_分面积段单价.png")

# ---- 图4 散点
fig, ax = plt.subplots(figsize=(8.5, 4.6))
palette = {"60㎡以下": GREY, "75-77㎡": C5, "94-97㎡": C3, "106-117㎡": C1, "124㎡": C4, "150㎡以上": "#4BACC6"}
for b, sub in df.groupby("bin"):
    ax.scatter(sub.date, sub.unit, s=26, color=palette[b], label=b, alpha=0.85, edgecolor="none")
comp = df[df.comp]
ax.scatter(comp.date, comp.unit, s=90, facecolor="none", edgecolor=C2, lw=1.4, label="本案可比样本(3室105-117㎡, 2026-03起)")
ax.axhline(12178, color=C2, ls="--", lw=1.2)
ax.text(pd.Timestamp("2025-07-25"), 12330, "本案估值单价 ≈ 12,178元/㎡（中性）", fontsize=8.5, color=C2)
ax.annotate("109.99㎡ 高楼层\n100.5万 (9,137)", (pd.Timestamp("2026-03-15"), 9137), xytext=(pd.Timestamp("2025-10-25"), 8300),
            fontsize=7.5, arrowprops=dict(arrowstyle="->", color=GREY), color=GREY)
ax.annotate("109.99㎡ 中楼层\n140万 (12,728)", (pd.Timestamp("2026-08-28"), 12728), xytext=(pd.Timestamp("2026-05-10"), 15600),
            fontsize=7.5, arrowprops=dict(arrowstyle="->", color=GREY), color=GREY)
ax.set_ylim(6500, 16500)
ax.set_ylabel("成交单价（元/㎡）")
ax.set_title("图4  80套成交逐套分布：同户型价差可达40万，怎么卖比何时卖更重要")
ax.legend(fontsize=7.5, loc="lower left", ncol=2)
ax.grid(alpha=0.3)
save(fig, "04_逐套成交散点.png")

# ---- 图5 武汉库存
fig, (a1, a2) = plt.subplots(1, 2, figsize=(9, 3.8))
a1.bar(["2025年中高点", "2026年6月末", "2026年7月"], [18.0, 15.7, 13.41], color=[C4, C1, C1])
for i, v in enumerate([18.0, 15.7, 13.41]):
    a1.text(i, v + 0.2, f"约{v:.1f}万套", ha="center", fontsize=9)
a1.set_ylim(0, 21); a1.set_title("武汉二手房挂牌量（万套）", fontsize=10)
a1.text(0.02, 0.02, "注：18万为媒体口径高点；15.7万按易居“较高点降13%”推算；\n13.41万为湖北中原2026年7月数据", transform=a1.transAxes, fontsize=6.5, color=GREY)
a2.bar(["2024上半年", "2025年1月末", "2026年8月末"], [30, 24, 26.2], color=[C4, C1, C1])
for i, v in enumerate([30, 24, 26.2]):
    a2.text(i, v + 0.5, f"{v:.0f}个月" if i == 0 else f"{v:.1f}个月", ha="center", fontsize=9)
a2.axhline(18, color=C2, ls="--", lw=1); a2.text(2.4, 18.5, "18个月=合理线", fontsize=7.5, color=C2, ha="right")
a2.set_ylim(0, 36); a2.set_title("武汉新房狭义库存去化周期", fontsize=10)
a2.text(0.02, 0.02, "注：2024上半年为“超30个月”，按30计；数据来源克而瑞/机构口径", transform=a2.transAxes, fontsize=6.5, color=GREY)
fig.suptitle("图5  武汉库存：二手挂牌明显消化，新房去化仍需两年多", fontsize=11)
save(fig, "05_武汉库存.png")

# ---- 图6 武汉二手房价同比
fig, ax = plt.subplots(figsize=(8, 3.8))
labels = ["2026-01\n武汉(中指)", "2026-03\n武汉(中指)", "2026-07\n武汉(中指)", "2026-08\n全国二线(统计局)", "2026-08\n一线(统计局)"]
vals = [-13.14, -12.66, -10.40, -4.9, -2.7]
bars = ax.bar(labels, vals, color=[C2, C2, C2, C1, C3])
for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width()/2, v - 0.6, f"{v:.1f}%", ha="center", fontsize=9, va="top")
ax.axhline(0, color="black", lw=0.8)
ax.set_ylim(-16, 2)
ax.set_ylabel("二手住宅价格同比（%）")
ax.set_title("图6  武汉二手房价同比跌幅在收窄，但仍是两位数，明显弱于一线")
ax.grid(axis="y", alpha=0.3)
save(fig, "06_武汉二手房价同比.png")

# ---- 图7 全国房地产指标
fig, ax = plt.subplots(figsize=(8, 3.8))
labels = ["开发投资", "新开工面积", "竣工面积", "新房销售面积", "新房销售额", "到位资金"]
vals = [-18.0, -23.4, -23.7, -11.6, -13.6, -20.2]
bars = ax.barh(labels, vals, color=C1)
for b, v in zip(bars, vals):
    ax.text(v - 0.5, b.get_y() + b.get_height()/2, f"{v:.1f}%", va="center", ha="right", fontsize=9)
ax.set_xlim(-28, 2)
ax.set_xlabel("2026年1–6月 同比（%）")
ax.set_title("图7  全国房地产供给端仍在深度收缩（供应减少是未来的利好）")
ax.grid(axis="x", alpha=0.3)
save(fig, "07_全国房地产指标.png")

# ---- 图8 全国出生人口
fig, ax = plt.subplots(figsize=(8, 3.8))
yrs = ["2021", "2022", "2023", "2024", "2025"]
births = [1062, 956, 902, 954, 792]
bars = ax.bar(yrs, births, color=[C1, C1, C1, C1, C2])
for b, v in zip(bars, births):
    ax.text(b.get_x() + b.get_width()/2, v + 12, f"{v}万", ha="center", fontsize=9)
ax.set_ylim(0, 1200)
ax.set_ylabel("出生人口（万人）")
ax.set_title("图8  全国出生人口：2025年跌破800万，总人口连续4年减少")
ax.grid(axis="y", alpha=0.3)
save(fig, "08_全国出生人口.png")

# ---- 图9 武汉常住人口
fig, ax = plt.subplots(figsize=(8, 3.8))
yrs = ["2021", "2022", "2023", "2024", "2025"]
pop = [1364.89, 1373.90, 1377.40, 1380.91, 1386.19]
inc = ["", "+9.0万", "+3.5万", "+3.5万", "+5.3万"]
ax.plot(yrs, pop, marker="o", color=C1, lw=2.2)
for x, y, i_ in zip(yrs, pop, inc):
    ax.annotate(f"{y:.1f}", (x, y), textcoords="offset points", xytext=(0, 8), ha="center", fontsize=8.5)
    if i_:
        ax.annotate(i_, (x, y), textcoords="offset points", xytext=(0, -16), ha="center", fontsize=8, color=C2)
ax.set_ylim(1355, 1395)
ax.set_ylabel("常住人口（万人）")
ax.set_title("图9  武汉常住人口连续增长，但增量已从两位数降到每年3–5万")
ax.text(0.02, 0.05, "2025年增量5.28万（+0.38%），为近三年最高；同期湖北全省减少23万，17个市州中只有武汉正增长", transform=ax.transAxes, fontsize=8, color=GREY)
ax.grid(axis="y", alpha=0.3)
save(fig, "09_武汉常住人口.png")

# ---- 图10 情景路径
V0 = 134.0
years = np.arange(0, 6)
scen = {"乐观（年+2.5%）": 0.025, "基准（年-2.5%）": -0.025, "悲观（年-6%）": -0.06}
fig, ax = plt.subplots(figsize=(8, 4))
for (k, gr), c in zip(scen.items(), [C3, C1, C2]):
    path = V0 * (1 + gr) ** years
    ax.plot(2026 + years, path, marker="o", lw=2.2, color=c, label=k)
    ax.annotate(f"{path[-1]:.0f}万", (2031, path[-1]), textcoords="offset points", xytext=(6, 0), fontsize=8.5, color=c)
ax.axhline(V0, color=GREY, ls=":", lw=1)
ax.set_xlabel("年份"); ax.set_ylabel("本案总价（万元）")
ax.set_title("图10  本案（110㎡）未来5年三种情景的价格路径")
ax.legend(fontsize=9)
ax.grid(alpha=0.3)
save(fig, "10_情景路径.png")

# ---- 图11 持有-出售差额
d, c, N, R, r = 0.02, 0.01, 3, 2.6304, 0.02
gs = np.linspace(-0.08, 0.06, 57)
sell = V0 * (1 - d) * (1 - c) * (1 + r) ** N
hold = V0 * (1 + gs) ** N * (1 - d) * (1 - c) + R * N
diff = hold - sell
fig, ax = plt.subplots(figsize=(8, 4))
ax.plot(gs * 100, diff, color=C1, lw=2.4)
ax.fill_between(gs * 100, diff, 0, where=diff < 0, color=C2, alpha=0.15, label="出售占优")
ax.fill_between(gs * 100, diff, 0, where=diff >= 0, color=C3, alpha=0.2, label="持有占优")
ax.axhline(0, color="black", lw=0.8)
be = ((1 + r) ** N - R * N / (V0 * (1 - d) * (1 - c))) ** (1 / N) - 1
ax.axvline(be * 100, color=C4, ls="--")
ax.text(be * 100 + 0.2, diff.max() * 0.85, f"盈亏平衡：房价年变动 ≈ {be*100:.1f}%", fontsize=9, color=C4)
ax.set_xlabel("未来3年房价年均变动率（%）")
ax.set_ylabel("持有 − 出售 的3年后资产差额（万元）")
ax.set_title("图11  只要房价不涨，现在卖出就不吃亏（3年期，租金2.0%净回报，资金收益2%）")
ax.legend(fontsize=9, loc="upper left")
ax.grid(alpha=0.3)
save(fig, "11_持有vs出售.png")

# ---- 图12 利率与回报
fig, ax = plt.subplots(figsize=(8, 3.6))
labels = ["5年期以上LPR", "新发房贷利率(6月)", "公积金首套(5年以上)", "本案净租金回报(估)", "存款/理财参考(假设)"]
vals = [3.5, 3.1, 2.6, 2.0, 1.5]
bars = ax.bar(labels, vals, color=[C1, C1, C1, C2, GREY])
for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width()/2, v + 0.05, f"{v:.1f}%", ha="center", fontsize=9)
ax.set_ylim(0, 4.3)
ax.set_ylabel("%")
ax.set_title("图12  租金回报低于房贷利率：持有的“显性成本”高于“显性收益”")
plt.setp(ax.get_xticklabels(), fontsize=8)
ax.grid(axis="y", alpha=0.3)
save(fig, "12_利率与回报.png")

print("charts done:", sorted(os.listdir(OUT)))
print("breakeven", be)

# ---- 图0 分析框架
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
fig, ax = plt.subplots(figsize=(9, 3.6))
ax.set_xlim(0, 10); ax.set_ylim(0, 4); ax.axis("off")
boxes = [
    (0.3, "微观\n小区80套成交\n价格·量·估值", C1),
    (2.7, "中观\n武汉楼市\n量·价·库存·结构", C5),
    (5.1, "宏观 & 长期\n经济·利率\n人口·政策", C3),
    (7.5, "推演 & 算账\n三情景\n持有 vs 出售", C2),
]
for x, t, c in boxes:
    ax.add_patch(FancyBboxPatch((x, 1.0), 2.2, 2.2, boxstyle="round,pad=0.08", fc=c, ec="none", alpha=0.92))
    ax.text(x + 1.1, 2.1, t, ha="center", va="center", color="white", fontsize=10.5, fontweight="bold", linespacing=1.5)
for x in (2.5, 4.9, 7.3):
    ax.add_patch(FancyArrowPatch((x + 0.02, 2.1), (x + 0.2, 2.1), arrowstyle="-|>", mutation_scale=18, color=GREY, lw=2))
ax.text(5, 0.45, "结论：先让小区自己的数据说话，再看大环境会不会改变它，最后把两条路的账算清楚",
        ha="center", fontsize=9.5, color=GREY)
save(fig, "00_分析框架.png")
