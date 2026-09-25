```python
import math
import re
from datetime import date, datetime

import pandas as pd
import plotly.express as px
import streamlit as st


# ============================================================
# PAGE
# ============================================================

st.set_page_config(
    page_title="창고별·일자별 입고 예정 PLT Dashboard",
    page_icon="📦",
    layout="wide",
)


# ============================================================
# SETTINGS
# ============================================================

EXCLUDE_DATE_WORDS = [
    "협의중",
    "합계",
    "total",
    "누계",
    "s&op",
    "sop",
]


# ============================================================
# BASIC FUNCTIONS
# ============================================================

def clean_col_name(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def to_number(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(" ", "", regex=False)
        .replace({
            "": None,
            "nan": None,
            "None": None,
        }),
        errors="coerce",
    )


def find_column(columns, candidates):
    columns = list(columns)

    # 정확히 일치
    for candidate in candidates:
        candidate = clean_col_name(candidate)

        for column in columns:
            if clean_col_name(column) == candidate:
                return column

    # 부분 일치
    for candidate in candidates:
        candidate = clean_col_name(candidate)

        for column in columns:
            if candidate and candidate in clean_col_name(column):
                return column

    return None


# ============================================================
# DATE DETECTION
# ============================================================

def parse_date_header(value, target_year):
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(value).normalize()

    value = normalize_text(value)

    if not value:
        return None

    if any(
        word.lower() in value.lower()
        for word in EXCLUDE_DATE_WORDS
    ):
        return None

    # 2026-09-01
    match = re.fullmatch(
        r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})",
        value,
    )

    if match:
        year, month, day = map(int, match.groups())

        try:
            return pd.Timestamp(year, month, day)
        except ValueError:
            return None

    # 9월 1일
    match = re.fullmatch(
        r"(\d{1,2})\s*월\s*(\d{1,2})\s*일?",
        value,
    )

    if match:
        month, day = map(int, match.groups())

        try:
            return pd.Timestamp(target_year, month, day)
        except ValueError:
            return None

    # 9. 1 / 09.01 / 9/1
    match = re.fullmatch(
        r"(\d{1,2})\s*[./]\s*(\d{1,2})",
        value,
    )

    if match:
        month, day = map(int, match.groups())

        try:
            return pd.Timestamp(target_year, month, day)
        except ValueError:
            return None

    return None


def detect_date_columns(
    df,
    target_year,
    target_month,
):
    detected = []

    for column in df.columns:
        parsed = parse_date_header(
            column,
            target_year,
        )

        if parsed is not None:
            if (
                parsed.year == target_year
                and parsed.month == target_month
            ):
                detected.append(
                    (column, parsed)
                )

    detected.sort(
        key=lambda x: x[1]
    )

    return detected


# ============================================================
# WAREHOUSE MAPPING
# ============================================================

def determine_warehouse(
    category,
    product_name,
):
    category = normalize_text(category)
    product_name = normalize_text(product_name)

    # 국내
    if category == "국내":
        return "부발4층"

    # 그레이스
    if category == "그레이스":
        return "그레이스"

    # 대만
    if category == "대만":
        return "대만"

    # 세포라
    if category == "세포라":
        return "세포라"

    # 슬리브X
    if category == "슬리브X":
        return "슬리브X"

    # 일본
    if category == "일본":
        return "어크로스비"

    # 직납
    if category == "직납":
        return "직납"

    # 캐나다
    if category == "캐나다":
        return "영문 부발4층"

    # 글로벌
    if category == "글로벌":

        if (
            "EF" in product_name
            or "글로벌" in product_name
        ):
            return "안성개정"

        if (
            "EU" in product_name
            or "유럽" in product_name
        ):
            return "부발1,2층"

    # 어크로스
    if category == "어크로스":

        if (
            "일문" in product_name
            or "JP" in product_name
        ):
            return "어크로스비"

    # 영문
    if category == "영문":

        if (
            "영문" in product_name
            or "글로벌" in product_name
        ):
            return "안성개성"

    # 유럽/영국
    if category == "유럽/영국":

        if (
            "EU" in product_name
            or "유럽재고" in product_name
        ):
            return "부발1,2층"

    # 아르고 계열
    if category.startswith("아르고"):
        return category

    # 미매핑
    return "미매핑"


# ============================================================
# PLT CALCULATION
# ============================================================

def calculate_plt(
    pcs,
    pallet_qty,
    round_partial=True,
):
    if pd.isna(pcs):
        return 0

    if pd.isna(pallet_qty):
        return 0

    try:
        pcs = float(pcs)
        pallet_qty = float(pallet_qty)
    except (TypeError, ValueError):
        return 0

    if pcs <= 0:
        return 0

    if pallet_qty <= 0:
        return 0

    result = pcs / pallet_qty

    if round_partial:
        return math.ceil(result)

    return result


# ============================================================
# FILE READ
# ============================================================

def read_uploaded_file(uploaded_file):

    filename = uploaded_file.name.lower()

    if filename.endswith(
        (".xlsx", ".xls")
    ):
        return pd.read_excel(
            uploaded_file,
            sheet_name=None,
            dtype=object,
        )

    if filename.endswith(".csv"):

        raw = uploaded_file.getvalue()

        from io import BytesIO

        for encoding in [
            "utf-8-sig",
            "cp949",
            "euc-kr",
            "utf-8",
        ]:
            try:
                return {
                    "CSV": pd.read_csv(
                        BytesIO(raw),
                        encoding=encoding,
                        dtype=object,
                    )
                }
            except UnicodeDecodeError:
                continue

        raise ValueError(
            "CSV 파일의 인코딩을 읽을 수 없습니다."
        )

    raise ValueError(
        "xlsx, xls, csv 파일만 지원합니다."
    )


# ============================================================
# SOURCE COLUMN DETECTION
# ============================================================

def detect_source_columns(df):

    columns = list(df.columns)

    return {
        "sku": find_column(
            columns,
            [
                "SKU ID",
                "SKU",
                "품목코드",
                "상품코드",
            ],
        ),

        "product": find_column(
            columns,
            [
                "제품 모델명",
                "제품명",
                "상품명",
                "상품명칭",
            ],
        ),

        "barcode": find_column(
            columns,
            [
                "바코드",
                "Barcode",
                "BARCODE",
            ],
        ),

        "category": find_column(
            columns,
            [
                "라인",
                "구분",
                "카테고리",
            ],
        ),

        "pallet_qty": find_column(
            columns,
            [
                "팔레트당 제품 수량",
                "팔레트 당 제품 수량",
                "PLT당 PCS 적재 수량",
                "팔레트당 PCS",
            ],
        ),
    }


# ============================================================
# WIDE → LONG
# ============================================================

def transform_inbound_data(
    df,
    date_columns,
    mapping,
    round_partial=True,
):

    required = [
        "sku",
        "product",
        "category",
        "pallet_qty",
    ]

    missing = [
        key
        for key in required
        if mapping.get(key) is None
    ]

    if missing:
        raise ValueError(
            "필수 컬럼을 찾지 못했습니다: "
            + ", ".join(missing)
        )

    sku_col = mapping["sku"]
    product_col = mapping["product"]
    barcode_col = mapping["barcode"]
    category_col = mapping["category"]
    pallet_col = mapping["pallet_qty"]

    base = pd.DataFrame()

    base["SKU ID"] = (
        df[sku_col]
        .map(normalize_text)
    )

    base["제품명"] = (
        df[product_col]
        .map(normalize_text)
    )

    if barcode_col:
        base["바코드"] = (
            df[barcode_col]
            .map(normalize_text)
        )
    else:
        base["바코드"] = ""

    base["구분"] = (
        df[category_col]
        .map(normalize_text)
    )

    base["팔레트당 제품 수량"] = (
        to_number(
            df[pallet_col]
        )
    )

    source_date_columns = [
        column
        for column, _ in date_columns
    ]

    for column in source_date_columns:
        base[column] = to_number(
            df[column]
        )

    # 가로형 날짜 데이터를 세로형으로 변환
    inbound_data = base.melt(
        id_vars=[
            "SKU ID",
            "제품명",
            "바코드",
            "구분",
            "팔레트당 제품 수량",
        ],
        value_vars=source_date_columns,
        var_name="입고일자_원본",
        value_name="입고 PCS",
    )

    # 날짜 매핑
    date_map = {
        column: parsed
        for column, parsed in date_columns
    }

    inbound_data["입고일자"] = (
        inbound_data[
            "입고일자_원본"
        ].map(date_map)
    )

    # PCS 숫자화
    inbound_data["입고 PCS"] = (
        pd.to_numeric(
            inbound_data["입고 PCS"],
            errors="coerce",
        )
        .fillna(0)
    )

    # 0 PCS 제거
    inbound_data = inbound_data[
        inbound_data["입고 PCS"] > 0
    ].copy()

    # 창고 매핑
    inbound_data["창고"] = (
        inbound_data.apply(
            lambda row:
                determine_warehouse(
                    row["구분"],
                    row["제품명"],
                ),
            axis=1,
        )
    )

    # PLT 계산
    inbound_data["입고 예정 PLT"] = (
        inbound_data.apply(
            lambda row:
                calculate_plt(
                    row["입고 PCS"],
                    row[
                        "팔레트당 제품 수량"
                    ],
                    round_partial,
                ),
            axis=1,
        )
    )

    inbound_data["입고일자"] = (
        pd.to_datetime(
            inbound_data["입고일자"],
            errors="coerce",
        )
    )

    return (
        inbound_data
        .sort_values(
            [
                "입고일자",
                "창고",
                "제품명",
            ]
        )
        .reset_index(drop=True)
    )


# ============================================================
# VALIDATION
# ============================================================

def make_validation(
    inbound_data,
):

    return {
        "전체 입고 PCS":
            inbound_data[
                "입고 PCS"
            ].sum(),

        "전체 입고 예정 PLT":
            inbound_data[
                "입고 예정 PLT"
            ].sum(),

        "SKU 수":
            inbound_data[
                "SKU ID"
            ].nunique(),

        "미매핑 건수":
            (
                inbound_data[
                    "창고"
                ] == "미매핑"
            ).sum(),

        "팔레트 수량 누락":
            (
                inbound_data[
                    "팔레트당 제품 수량"
                ].isna()
                |
                (
                    inbound_data[
                        "팔레트당 제품 수량"
                    ] <= 0
                )
            ).sum(),

        "입고 데이터 건수":
            len(inbound_data),
    }


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header("설정")

target_year = st.sidebar.number_input(
    "연도",
    min_value=2020,
    max_value=2100,
    value=2026,
    step=1,
)

target_month = st.sidebar.selectbox(
    "대상 월",
    list(range(1, 13)),
    index=8,
    format_func=lambda x:
        f"{x}월",
)

round_partial = st.sidebar.checkbox(
    "부분 팔레트 올림",
    value=True,
)


# ============================================================
# TITLE
# ============================================================

st.title(
    "📦 창고별 · 일자별 입고 예정 PLT Dashboard"
)

st.caption(
    "PU팀 SCM 입고 예정 PCS → "
    "창고별 / 일자별 입고 예정 PLT 자동 환산"
)


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_file = st.file_uploader(
    "PU 입고 일정 파일 업로드",
    type=[
        "xlsx",
        "xls",
        "csv",
    ],
)

if uploaded_file is None:

    st.info(
        "PU 입고 일정 파일을 업로드하세요."
    )

    st.stop()


# ============================================================
# READ FILE
# ============================================================

try:

    sheets = read_uploaded_file(
        uploaded_file
    )

except Exception as e:

    st.error(
        f"파일을 읽을 수 없습니다: {e}"
    )

    st.stop()


# ============================================================
# SHEET
# ============================================================

sheet_name = st.selectbox(
    "사용할 시트",
    list(sheets.keys()),
)

raw_df = sheets[
    sheet_name
].copy()

if raw_df.empty:

    st.error(
        "선택한 시트에 데이터가 없습니다."
    )

    st.stop()


raw_df.columns = [
    clean_col_name(column)
    for column in raw_df.columns
]


# ============================================================
# COLUMN MAPPING
# ============================================================

st.subheader(
    "데이터 인식"
)

mapping = detect_source_columns(
    raw_df
)

mapping_display = {
    "SKU ID":
        mapping["sku"],

    "제품명":
        mapping["product"],

    "바코드":
        mapping["barcode"],

    "구분/라인":
        mapping["category"],

    "팔레트당 제품 수량":
        mapping["pallet_qty"],
}

st.dataframe(
    pd.DataFrame(
        {
            "항목":
                list(
                    mapping_display.keys()
                ),

            "원본 컬럼":
                list(
                    mapping_display.values()
                ),
        }
    ),
    hide_index=True,
    use_container_width=True,
)


# ============================================================
# DATE COLUMNS
# ============================================================

date_columns = detect_date_columns(
    raw_df,
    target_year,
    target_month,
)

if not date_columns:

    st.error(
        f"{target_year}년 "
        f"{target_month}월 날짜 컬럼을 "
        "찾지 못했습니다."
    )

    st.stop()


st.success(
    f"{len(date_columns)}개 날짜 컬럼 인식"
)


# ============================================================
# TRANSFORM
# ============================================================

try:

    inbound_data = (
        transform_inbound_data(
            raw_df,
            date_columns,
            mapping,
            round_partial,
        )
    )

except Exception as e:

    st.error(
        f"데이터 변환 중 오류가 발생했습니다: {e}"
    )

    st.stop()


if inbound_data.empty:

    st.warning(
        "선택한 월에 입고 예정 PCS가 없습니다."
    )

    st.stop()


# ============================================================
# VALIDATION
# ============================================================

validation = make_validation(
    inbound_data
)

st.subheader(
    "데이터 검증"
)

c1, c2, c3, c4 = st.columns(4)

c1.metric(
    "전체 입고 PCS",
    f"{validation['전체 입고 PCS']:,.0f}",
)

c2.metric(
    "전체 입고 예정 PLT",
    f"{validation['전체 입고 예정 PLT']:,.0f}",
)

c3.metric(
    "SKU 수",
    f"{validation['SKU 수']:,}",
)

c4.metric(
    "미매핑 건수",
    f"{validation['미매핑 건수']:,}",
)

if validation["미매핑 건수"] > 0:

    st.warning(
        "미매핑 데이터가 있습니다."
    )

if validation["팔레트 수량 누락"] > 0:

    st.warning(
        "팔레트당 제품 수량이 없거나 "
        "0인 데이터가 있습니다."
    )


# ============================================================
# FILTER
# ============================================================

st.subheader(
    "필터"
)

f1, f2 = st.columns(2)

warehouse_options = sorted(
    inbound_data[
        "창고"
    ]
    .dropna()
    .unique()
    .tolist()
)

category_options = sorted(
    inbound_data[
        "구분"
    ]
    .dropna()
    .unique()
    .tolist()
)

with f1:

    selected_warehouses = st.multiselect(
        "창고",
        warehouse_options,
        default=warehouse_options,
    )

with f2:

    selected_categories = st.multiselect(
        "구분/라인",
        category_options,
        default=category_options,
    )


filtered = inbound_data[
    inbound_data["창고"].isin(
        selected_warehouses
    )
    &
    inbound_data["구분"].isin(
        selected_categories
    )
].copy()


if filtered.empty:

    st.warning(
        "현재 필터 조건에 해당하는 "
        "데이터가 없습니다."
    )

    st.stop()


# ============================================================
# KPI
# ============================================================

daily_plt = (
    filtered
    .groupby(
        "입고일자",
        as_index=False,
    )[
        "입고 예정 PLT"
    ]
    .sum()
    .sort_values(
        "입고일자"
    )
)

total_plt = (
    filtered[
        "입고 예정 PLT"
    ].sum()
)

total_pcs = (
    filtered[
        "입고 PCS"
    ].sum()
)

sku_count = (
    filtered[
        "SKU ID"
    ].nunique()
)

max_daily_plt = (
    daily_plt[
        "입고 예정 PLT"
    ].max()
)


k1, k2, k3, k4 = st.columns(4)

k1.metric(
    "입고 예정 PLT",
    f"{total_plt:,.0f}",
)

k2.metric(
    "입고 예정 PCS",
    f"{total_pcs:,.0f}",
)

k3.metric(
    "SKU",
    f"{sku_count:,}",
)

k4.metric(
    "일 최대 입고 PLT",
    f"{max_daily_plt:,.0f}",
)


# ============================================================
# CHART 1
# 일자별 전체 입고 예정 PLT 추이
# ============================================================

st.subheader(
    "1. 일자별 전체 입고 예정 PLT 추이"
)

fig1 = px.line(
    daily_plt,
    x="입고일자",
    y="입고 예정 PLT",
    markers=True,
)

fig1.update_layout(
    xaxis_title="입고일자",
    yaxis_title="입고 예정 PLT",
    hovermode="x unified",
    height=420,
)

st.plotly_chart(
    fig1,
    use_container_width=True,
)


# ============================================================
# CHART 2
# 창고별 일자별 입고 예정 PLT
# ============================================================

st.subheader(
    "2. 창고별 일자별 입고 예정 PLT"
)

warehouse_daily = (
    filtered
    .groupby(
        [
            "입고일자",
            "창고",
        ],
        as_index=False,
    )[
        "입고 예정 PLT"
    ]
    .sum()
    .sort_values(
        "입고일자"
    )
)

fig2 = px.bar(
    warehouse_daily,
    x="입고일자",
    y="입고 예정 PLT",
    color="창고",
    barmode="stack",
)

fig2.update_layout(
    xaxis_title="입고일자",
    yaxis_title="입고 예정 PLT",
    height=500,
)

st.plotly_chart(
    fig2,
    use_container_width=True,
)


# ============================================================
# CHART 3
# 창고별 총 입고 예정 PLT 구성
# ============================================================

st.subheader(
    "3. 창고별 총 입고 예정 PLT 구성"
)

warehouse_total = (
    filtered
    .groupby(
        "창고",
        as_index=False,
    )[
        "입고 예정 PLT"
    ]
    .sum()
    .sort_values(
        "입고 예정 PLT",
        ascending=True,
    )
)

fig3 = px.bar(
    warehouse_total,
    x="입고 예정 PLT",
    y="창고",
    orientation="h",
    text="입고 예정 PLT",
)

fig3.update_layout(
    xaxis_title="입고 예정 PLT",
    yaxis_title="창고",
    height=450,
)

st.plotly_chart(
    fig3,
    use_container_width=True,
)


# ============================================================
# CHART 4
# 품목별 입고 예정 PLT TOP 10
# ============================================================

st.subheader(
    "4. 품목별 입고 예정 PLT TOP 10"
)

sku_top10 = (
    filtered
    .groupby(
        [
            "SKU ID",
            "제품명",
        ],
        as_index=False,
    )[
        "입고 예정 PLT"
    ]
    .sum()
    .sort_values(
        "입고 예정 PLT",
        ascending=False,
    )
    .head(10)
)

sku_top10["품목"] = (
    sku_top10["제품명"]
    + " ("
    + sku_top10["SKU ID"]
    + ")"
)

fig4 = px.bar(
    sku_top10.sort_values(
        "입고 예정 PLT"
    ),
    x="입고 예정 PLT",
    y="품목",
    orientation="h",
    text="입고 예정 PLT",
)

fig4.update_layout(
    xaxis_title="입고 예정 PLT",
    yaxis_title="품목",
    height=500,
)

st.plotly_chart(
    fig4,
    use_container_width=True,
)


# ============================================================
# CHART 5
# 일자별 창고 입고 집중도
# ============================================================

st.subheader(
    "5. 일자별 창고 입고 집중도"
)

heatmap_data = (
    filtered
    .pivot_table(
        index="입고일자",
        columns="창고",
        values="입고 예정 PLT",
        aggfunc="sum",
        fill_value=0,
    )
    .sort_index()
)

fig5 = px.imshow(
    heatmap_data,
    aspect="auto",
    labels={
        "x": "창고",
        "y": "입고일자",
        "color": "입고 예정 PLT",
    },
)

fig5.update_layout(
    height=600,
)

st.plotly_chart(
    fig5,
    use_container_width=True,
)


# ============================================================
# UNMAPPED DATA
# ============================================================

st.subheader(
    "창고 미매핑 데이터"
)

unmapped = (
    inbound_data[
        inbound_data["창고"]
        == "미매핑"
    ]
    .groupby(
        [
            "구분",
            "제품명",
            "SKU ID",
        ],
        as_index=False,
    )[
        "입고 PCS"
    ]
    .sum()
    .sort_values(
        "입고 PCS",
        ascending=False,
    )
)

if unmapped.empty:

    st.success(
        "미매핑 데이터가 없습니다."
    )

else:

    st.dataframe(
        unmapped,
        hide_index=True,
        use_container_width=True,
    )


# ============================================================
# FINAL DATA
# ============================================================

st.subheader(
    "변환된 입고 DATA"
)

display_columns = [
    "입고일자",
    "창고",
    "구분",
    "SKU ID",
    "제품명",
    "바코드",
    "입고 PCS",
    "팔레트당 제품 수량",
    "입고 예정 PLT",
]

display_df = filtered[
    display_columns
].copy()

st.dataframe(
    display_df,
    hide_index=True,
    use_container_width=True,
    height=500,
)


# ============================================================
# DOWNLOAD
# ============================================================

csv_data = display_df.to_csv(
    index=False,
    encoding="utf-8-sig",
)

st.download_button(
    label="📥 변환된 입고 DATA CSV 다운로드",
    data=csv_data,
    file_name=(
        f"inbound_plt_"
        f"{target_year}_"
        f"{target_month:02d}.csv"
    ),
    mime="text/csv",
)
```
