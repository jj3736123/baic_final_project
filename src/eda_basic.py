"""
Air France 검색광고 1차 EDA (기본 집계)
- 데이터: AirFrance Search data.xlsx 의 숨김 시트 'DoubleClick Data' (4,510행)
- 분류 기준: docs/keyword_classification_tree.svg, 노션 EDA·전처리 > 키워드 분류 기준
- 실행: python src/eda_basic.py  (data/ 폴더에 원본 엑셀을 넣고 실행)
"""
import re
import sys
import pandas as pd

from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
PATH = sys.argv[1] if len(sys.argv) > 1 else ROOT / "data" / "AirFrance Search data.xlsx"
OUT = ROOT / "outputs" / "eda_basic_results.xlsx"
OUT.parent.mkdir(exist_ok=True)
df = pd.read_excel(PATH, sheet_name="DoubleClick Data")
df = df.rename(columns={"Total Volume of Bookings": "Bookings"})
df["kw"] = df["Keyword"].str.lower().str.strip()

# ── 1. 전처리 ────────────────────────────────────────────
# 검색광고가 아닌 콘텐츠 네트워크 광고(cont:) 48행 제외
df = df[~df["kw"].str.startswith("cont:")].copy()

# 매체 그룹: Overture는 Yahoo의 검색광고 플랫폼이라 엔진 단위 비교 시 함께 묶음
df["engine"] = df["Publisher Name"].str.split(" - ").str[0].replace({"Overture": "Yahoo"})

# ── 2. 분류 변수 ─────────────────────────────────────────
# 브랜드 여부
df["brand"] = df["kw"].str.contains(r"air\s*france|airfrance|airframce", regex=True)
df["brand_label"] = df["brand"].map({True: "브랜드", False: "비브랜드"})

# broad / focused (비브랜드만): 국가·도시·공항명 포함 여부
COUNTRIES = "france spain italy turkey greece india germany uk england britain switzerland tahiti".split()
CITIES = """paris rome athens nice barcelona venice madrid dubai bordeaux lyon frankfurt prague florence marseille
marseilles delhi cairo milan munich berlin london marrakech johannesburg moscow split nairobi sofia cologne amsterdam
geneva kiev istanbul tunis amman belgrade casablanca hamburg stuttgart budapest mumbai dakar zurich brussels toulouse
bora agadir hanover copenhagen nuremberg warsow warsaw damascus strasbourg orlando vienna lourdes bucarest bucharest
naples montpellier cannes rabat dusseldorf bangalore abidjan pisa chennai nouakchott helsinki petersburg conakry
gothenburg bologna lisbon nimes oslo dijon manchester nantes seville djibouti cincinnati cotonou grenoble zagreb
perpignan malaga bamako stockholm krakow glasgow rochelle limoges ouagadougou rennes bergerac caen mulhouse lorient
ajaccio valencia genoa lille bangui turin reims toulon poitiers carcassonne brest avignon yaound southampton basel
clermont rodez lom deauville dhabi bastia beauvais agen calais athenes aviv""".split()
ORIGINS = "york angeles francisco chicago miami atlanta detroit boston philadelphia houston washington dc ny nyc newark jfk lax diego phoenix".split()
AIRPORTS = "cdg orly roissy gaulle degaulle aeroport".split()
PLACES = set(COUNTRIES + CITIES + ORIGINS + AIRPORTS)

def has_place(k):
    return bool(set(re.findall(r"[a-z]+", k)) & PLACES)

nb = ~df["brand"]
df.loc[nb, "scope"] = df.loc[nb, "kw"].apply(lambda k: "focused" if has_place(k) else "broad")

# 검색 의도 (비브랜드만)
FLIGHT = r"flight|fly|fare|airline|air line|ticket|tiket|plane|air\b|airway"
TRAVEL = r"travel|vacation|trip|tour|holiday|getaway|romantic|romance|hotel|cruise|visit|tourism"

def intent(k):
    if "tour de france" in k:
        return "무관"
    if re.search(FLIGHT, k) or re.search(r"\w+ to \w+", k):  # 노선 형태(A to B)는 항공권
        return "항공권"
    if re.search(TRAVEL, k):
        return "여행"
    return "기타"

df.loc[nb, "intent"] = df.loc[nb, "kw"].apply(intent)

# ── 3. 집계 함수 ─────────────────────────────────────────
# 행별 비율을 평균내지 않고, 원천 수치를 합산한 뒤 비율을 다시 계산
def summarize(data, by):
    g = data.groupby(by, dropna=False).agg(
        keywords=("kw", "size"),
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

results = {
    "1_전체": summarize(df.assign(all="전체"), "all"),
    "2_매체별": summarize(df, "Publisher Name"),
    "3_검색엔진별": summarize(df, "engine"),
    "4_브랜드여부": summarize(df, "brand_label"),
    "5_broad_focused": summarize(df[nb], "scope"),
    "6_검색의도": summarize(df[nb], "intent"),
    "7_캠페인별": summarize(df, "Campaign"),
}

# 예약 0건 키워드
df["zero_booking"] = df["Bookings"] == 0
results["8_예약0건"] = summarize(df, "zero_booking")

# ── 4. 출력 ──────────────────────────────────────────────
fmt = {"cost_share": "{:.1%}", "CTR": "{:.2%}", "conv": "{:.2%}", "CPC": "${:.2f}", "CPA": "${:.0f}",
       "cost": "${:,.0f}", "revenue": "${:,.0f}", "net_revenue": "${:,.0f}", "ROA": "{:.1f}"}
with pd.ExcelWriter(OUT) as w:
    for name, t in results.items():
        t.to_excel(w, sheet_name=name)
        show = t.copy()
        for c, f in fmt.items():
            show[c] = show[c].map(lambda v, f=f: f.format(v) if pd.notna(v) and v != float("inf") else "-")
        print(f"\n===== {name} =====")
        print(show.drop(columns=["impressions"]).to_string())
