import requests
import pandas as pd
import plotly.express as px

from datetime import date, timedelta, datetime
from zoneinfo import ZoneInfo


# 최근 90일 환율 가져오기
end_date = date.today()
start_date = end_date - timedelta(days=90)

response = requests.get(
    "https://api.frankfurter.dev/v2/rates",
    params={
        "from": start_date.isoformat(),
        "base": "KRW",
        "quotes": "USD,JPY,EUR,CNY"
    },
    timeout=30
)

response.raise_for_status()
df = pd.DataFrame(response.json())


# 환율 데이터 정리
df["기준 단위"] = 1
df.loc[df["quote"] == "JPY", "기준 단위"] = 100
df["원화 환율"] = df["기준 단위"] / df["rate"]
df["date"] = pd.to_datetime(df["date"])

df = df.rename(columns={
    "date": "날짜",
    "quote": "통화"
})

df = df[
    ["날짜", "통화", "기준 단위", "원화 환율"]
].sort_values(["통화", "날짜"])


# 변화율과 변동성 계산
df["전일 변화율(%)"] = (
    df.groupby("통화")["원화 환율"]
    .pct_change(fill_method=None) * 100
)

df["1주 변화율(%)"] = (
    df.groupby("통화")["원화 환율"]
    .pct_change(5, fill_method=None) * 100
)

df["1개월 변화율(%)"] = (
    df.groupby("통화")["원화 환율"]
    .pct_change(20, fill_method=None) * 100
)

daily_change = (
    df.groupby("통화")["원화 환율"]
    .pct_change(fill_method=None)
)

df["20일 변동성(%)"] = (
    daily_change.groupby(df["통화"])
    .transform(lambda x: x.rolling(20).std() * (252 ** 0.5) * 100)
)


def risk_level(value):
    if pd.isna(value):
        return "데이터 부족"
    if value < 5:
        return "낮음"
    if value < 10:
        return "보통"
    return "높음"


df["환율 위험도"] = df["20일 변동성(%)"].apply(risk_level)

number_columns = [
    "원화 환율",
    "전일 변화율(%)",
    "1주 변화율(%)",
    "1개월 변화율(%)",
    "20일 변동성(%)"
]

df[number_columns] = df[number_columns].round(2)

latest_df = df.groupby("통화", as_index=False).tail(1).copy()


# 상대지수 계산
comparison_df = df.copy()

first_rate = (
    comparison_df.groupby("통화")["원화 환율"]
    .transform("first")
)

comparison_df["상대 지수"] = (
    comparison_df["원화 환율"] / first_rate * 100
).round(2)


# Excel 저장
with pd.ExcelWriter(
    "trade_fx_dashboard.xlsx",
    engine="openpyxl"
) as writer:

    latest_df.to_excel(
        writer,
        sheet_name="최신요약",
        index=False
    )

    df.to_excel(
        writer,
        sheet_name="90일전체",
        index=False
    )

    comparison_df.to_excel(
        writer,
        sheet_name="상대지수",
        index=False
    )


# 그래프 만들기
fig1 = px.line(
    df,
    x="날짜",
    y="원화 환율",
    color="통화",
    title="최근 90일 원화 환율"
)

fig1.update_layout(
    template="plotly_white",
    hovermode="x unified"
)

fig2 = px.line(
    comparison_df,
    x="날짜",
    y="상대 지수",
    color="통화",
    title="통화별 상대 변화"
)

fig2.add_hline(y=100, line_dash="dash")
fig2.update_layout(template="plotly_white")


# HTML용 표와 그래프
display_df = latest_df.copy()
display_df["날짜"] = display_df["날짜"].dt.strftime("%Y-%m-%d")

table_html = display_df.to_html(
    index=False,
    border=0,
    classes="summary-table"
)

graph1_html = fig1.to_html(
    full_html=False,
    include_plotlyjs="cdn"
)

graph2_html = fig2.to_html(
    full_html=False,
    include_plotlyjs=False
)

updated_at = datetime.now(
    ZoneInfo("Asia/Seoul")
).strftime("%Y-%m-%d %H:%M")


# 최종 홈페이지
html = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>무역 환율 리스크 대시보드</title>

<style>
body {{
    margin: 0;
    padding: 20px;
    background: #f3f5f9;
    color: #1f2937;
    font-family: "Apple SD Gothic Neo", Arial, sans-serif;
}}

.container {{
    max-width: 1200px;
    margin: auto;
}}

.card {{
    background: white;
    margin-bottom: 22px;
    padding: 20px;
    border-radius: 14px;
    overflow-x: auto;
    box-shadow: 0 3px 12px rgba(0,0,0,0.08);
}}

.summary-table {{
    width: 100%;
    border-collapse: collapse;
    white-space: nowrap;
}}

.summary-table th {{
    padding: 11px;
    color: white;
    background: #2563eb;
}}

.summary-table td {{
    padding: 11px;
    text-align: center;
    border-bottom: 1px solid #e5e7eb;
}}

.updated {{
    color: #6b7280;
}}
</style>
</head>

<body>
<div class="container">

<h1>무역 환율 리스크 대시보드</h1>
<p class="updated">최근 자동 갱신: {updated_at}</p>

<div class="card">
<h2>최신 환율 위험 요약</h2>
{table_html}
</div>

<div class="card">
{graph1_html}
</div>

<div class="card">
{graph2_html}
</div>

</div>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as file:
    file.write(html)

print("자동 갱신 완료:", updated_at)
