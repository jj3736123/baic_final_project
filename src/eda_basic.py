"""
Air France 검색광고 1차 EDA (기본 집계)
- 입력: data/airfrance_preprocessed.xlsx (src/preprocess.py 실행 결과)
- 출력: outputs/eda_basic_results.xlsx
- 실행: python src/preprocess.py → python src/eda_basic.py
"""
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
df = pd.read_excel(ROOT / "data" / "airfrance_preprocessed.xlsx", sheet_name="data")
df = df.rename(columns={"Total Volume of Bookings": "Bookings"})
OUT = ROOT / "outputs" / "eda_basic_results.xlsx"
OUT.parent.mkdir(exist_ok=True)

# 집계 원칙: 행별 비율을 평균내지 않고, 원천 수치를 합산한 뒤 비율을 다시 계산
def summarize(data, by):
    g = data.groupby(by, dropna=False).agg(
        keywords=("Keyword", "size"),
        impressions=("Impressions", "sum"),
        clicks=("Clicks", "sum"),
        cost=("Total Cost", "sum"),
        bookings=("Bookings", "sum"),
        revenue=("Amount", "sum"),
    )
    g["cost_share"] = g["cost"] / g["cost"].sum()
    g["CTR"] = g["clicks"] / g["impressions"]
    g["CPC"] = g["cost"] / g["clicks"]
    g["conv"] = g["bookings"] / g["clicks"]
    g["CPA"] = g["cost"] / g["bookings"]
    g["net_revenue"] = g["revenue"] - g["cost"]
    g["ROA"] = g["net_revenue"] / g["cost"]  # 데이터 컬럼과 동일 정의: 순수익 / 광고비
    return g.sort_values("cost", ascending=False)

nb = df[df["brand_label"] == "비브랜드"]
results = {
    "1_전체": summarize(df.assign(all="전체"), "all"),
    "2_매체별": summarize(df, "Publisher Name"),
    "3_검색엔진별": summarize(df, "engine"),
    "4_브랜드여부": summarize(df, "brand_label"),
    "5_매체x브랜드": summarize(df, ["Publisher Name", "brand_label"]),
    "6_broad_focused": summarize(nb, "scope"),
    "7_검색의도": summarize(nb, "intent"),
    "8_캠페인별": summarize(df, "campaign_group"),
    "9_예약0건": summarize(df.assign(zero=df["Bookings"] == 0), "zero"),
}

with pd.ExcelWriter(OUT) as w:
    for name, t in results.items():
        t.to_excel(w, sheet_name=name)
        print(f"\n===== {name} =====")
        print(t[["cost", "cost_share", "CPC", "conv", "CPA", "net_revenue", "ROA"]].round(3).to_string())
