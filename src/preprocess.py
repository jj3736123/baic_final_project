"""
Air France 검색광고 데이터 전처리

입력: data/AirFrance Search data.xlsx 의 숨김 시트 'DoubleClick Data' (4,510행)
출력: data/airfrance_preprocessed.xlsx  (시트: data, 컬럼설명)

전처리 내용 (팀 확정 사항)
1. Keyword가 'cont:'로 시작하는 48행 제외 (검색광고가 아닌 콘텐츠 네트워크 광고)
2. 컬럼 삭제: Click Charges(Total Cost와 동일), Keyword Type(전부 같은 값), Publisher ID(Publisher Name과 1:1),
   Category(대행사 내부 태그, 64%가 uncategorized), Probability of Booking(클릭률×전환율로 계산된 값)
3. 키워드 분류: 브랜드/비브랜드, 비브랜드 안에서 broad/focused, 검색 의도
4. 매체: Publisher Name 7개 그대로, 회사 단위 참고 컬럼(engine) 추가
5. 캠페인: Geo Targeted 13개를 하나로 통합, Unassigned는 '미지정'
6. Bid Strategy: 표기 통일, 결측은 '미설정'
7. 이상치: 행 삭제 없음. 키워드 단위 판단 대상 표시 컬럼 추가
   - flag_brand_cpc_100: 브랜드이면서 클릭 100회 이상 (브랜드 KPI인 클릭당 비용 비교용)
   - flag_conv_370: 비브랜드이면서 클릭 370회 이상 (전환율·ROA·Net Revenue 비교, 예약 0건 판단용)
"""
import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data" / "AirFrance Search data.xlsx"
OUT = ROOT / "data" / "airfrance_preprocessed.xlsx"

df = pd.read_excel(RAW, sheet_name="DoubleClick Data")
df = df.rename(columns={df.columns[0]: "Publisher ID"})
n_raw = len(df)

# ── 1. 콘텐츠 네트워크 광고 제외 ─────────────────────────
kw = df["Keyword"].str.lower().str.strip()
df = df[~kw.str.startswith("cont:")].copy()
df["kw"] = df["Keyword"].str.lower().str.strip()

# ── 2. 불필요한 컬럼 삭제 ───────────────────────────────
df = df.drop(columns=["Click Charges", "Keyword Type", "Publisher ID", "Category", "Probability of Booking"])

# ── 3. 키워드 분류 ──────────────────────────────────────
# 3-1. 브랜드 여부: air france / airfrance / 오타 airframce 포함
df["brand"] = df["kw"].str.contains(r"air\s*france|airfrance|airframce", regex=True)
df["brand_label"] = df["brand"].map({True: "브랜드", False: "비브랜드"})

# 3-2. broad / focused (비브랜드만): 국가·도시·공항명 포함 여부
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

def scope(k):
    return "focused" if set(re.findall(r"[a-z]+", k)) & PLACES else "broad"

# 3-3. 검색 의도 (비브랜드만)
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

nb = ~df["brand"]
df.loc[nb, "scope"] = df.loc[nb, "kw"].apply(scope)
df.loc[nb, "intent"] = df.loc[nb, "kw"].apply(intent)

# ── 4. 매체: 회사 단위 참고 컬럼 ─────────────────────────
df["engine"] = df["Publisher Name"].str.split(" - ").str[0].replace({"Overture": "Yahoo"})

# ── 5. 캠페인 정리 ──────────────────────────────────────
df["campaign_group"] = df["Campaign"]
df.loc[df["Campaign"].str.startswith("Geo Targeted"), "campaign_group"] = "Geo Targeted"
df.loc[df["Campaign"] == "Unassigned", "campaign_group"] = "미지정"

# ── 6. Bid Strategy 표기 통일 ───────────────────────────
BID_MAP = {
    "Position 5-10 Bid Strategy": "Position 5-10",
    "Position 2-5 Bid Strategy": "Position 2-5",
    "Position 1-2 Target": "Position 1-2",
    "Position 1 -2 Target": "Position 1-2",
    "Position 1- 3": "Position 1-3",
    "Position 1-4 Bid Strategy": "Position 1-4",
    "Postiion 1-4 Bid Strategy": "Position 1-4",
    "Pos 3-6": "Position 3-6",
}
df["Bid Strategy"] = df["Bid Strategy"].map(BID_MAP).fillna("미설정")

# ── 7. 최소 클릭 기준 표시 ──────────────────────────────
df["flag_brand_cpc_100"] = df["brand"] & (df["Clicks"] >= 100)
df["flag_conv_370"] = (~df["brand"]) & (df["Clicks"] >= 370)

df = df.drop(columns=["kw", "brand"])

# ── 저장 ────────────────────────────────────────────────
DICT = pd.DataFrame([
    ("brand_label", "브랜드 / 비브랜드 (air france / airfrance / airframce 포함 시 브랜드)"),
    ("scope", "비브랜드만: focused(국가·도시·공항명 포함) / broad"),
    ("intent", "비브랜드만: 항공권 / 여행 / 기타 / 무관(tour de france)"),
    ("engine", "회사 단위 매체 (Google / Yahoo / MSN, Overture는 Yahoo에 포함). 참고용"),
    ("campaign_group", "Campaign 정리본: Geo Targeted 13개 통합, Unassigned는 미지정"),
    ("Bid Strategy", "표기 통일, 결측은 미설정"),
    ("flag_brand_cpc_100", "브랜드이면서 클릭 100회 이상. 브랜드 키워드의 클릭당 비용을 키워드 단위로 비교할 때 이 행만 사용"),
    ("flag_conv_370", "비브랜드이면서 클릭 370회 이상. 전환율·ROA·Net Revenue 비교와 예약 0건 판단을 키워드 단위로 할 때 이 행만 사용. 클릭당 비용은 이 행 안에서 ROA 원인 분해용으로만 봄"),
    ("(집계 원칙)", "여러 행을 묶어 지표를 낼 때는 행별 비율을 평균내지 말고 원천 수치(매출·광고비·클릭·예약)를 합산한 뒤 다시 계산"),
], columns=["컬럼", "설명"])

with pd.ExcelWriter(OUT) as w:
    df.to_excel(w, sheet_name="data", index=False)
    DICT.to_excel(w, sheet_name="컬럼설명", index=False)

print(f"원본 {n_raw}행 → 전처리 후 {len(df)}행, {df.shape[1]}컬럼 → {OUT.relative_to(ROOT)}")
print(df["brand_label"].value_counts().to_dict()); print(list(df.columns))
print(df["scope"].value_counts().to_dict(), df["intent"].value_counts().to_dict())
print(df["campaign_group"].nunique(), "캠페인 그룹 |", df["Bid Strategy"].value_counts().to_dict())
print("flag_brand_cpc_100", int(df.flag_brand_cpc_100.sum()), "| flag_conv_370", int(df.flag_conv_370.sum()))
