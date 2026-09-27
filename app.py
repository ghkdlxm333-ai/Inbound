from datetime import datetime, timedelta
from io import BytesIO
import math
import re
import pandas as pd
import streamlit as st

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="올리브영 자동 출고 & LOT 매핑 (그레이스 3PL)",
    page_icon="📦",
    layout="wide",
)

MIN_SHELF_LIFE_DAYS = 547  # 올리브영 납품 기준 잔여 유통기한 (1년 6개월 = 547일)

# ============================================================
# DATA CLEANING HELPERS
# ============================================================
def clean_str(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str

def parse_number(val):
    try:
        if pd.isna(val):
            return 0
        cleaned = re.sub(r"[^\d.-]", "", str(val))
        return float(cleaned) if cleaned else 0
    except Exception:
        return 0

# ============================================================
# AUTOMATIC FILE PARSER (.xls, .xlsx, HTML, CSV 모두 자동 대응)
# ============================================================
def load_uploaded_file(uploaded_file):
    """
    WMS에서 받은 .xls 파일이 어떤 형태이든(바이너리, HTML-based 엑셀, CSV 등)
    사용자의 추가 작업 없이 자동으로 엔진을 맞춰 읽어옵니다.
    """
    raw_bytes = uploaded_file.getvalue()
    
    # 1. Excel (openpyxl / xlrd) 시도
    try:
        return pd.read_excel(BytesIO(raw_bytes), header=None)
    except Exception:
        pass

    # 2. HTML Table 형태의 .xls 엑셀 파일 시도
    try:
        html_dfs = pd.read_html(BytesIO(raw_bytes))
        if html_dfs:
            return html_dfs[0]
    except Exception:
        pass

    # 3. CSV 시도 (인코딩 자동 대조)
    for encoding in ["utf-8-sig", "cp949", "euc-kr", "utf-8", "latin1"]:
        try:
            return pd.read_csv(BytesIO(raw_bytes), encoding=encoding, header=None)
        except Exception:
            continue

    raise ValueError("파일을 정상적으로 읽을 수 없습니다. 파일 상태를 확인해주세요.")

# ============================================================
# WMS & DELIVERY FILE PARSERS
# ============================================================
def parse_grace_wms(df_raw):
    """
    그레이스 3PL WMS 재고 현황 파싱
    """
    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = " ".join(row.dropna().astype(str))
        if "상품코드" in row_str and "LOT" in row_str:
            header_idx = idx
            break

    if header_idx is None:
        header_idx = 1

    df = df_raw.iloc[header_idx + 1:].copy()
    
    parsed_rows = []
    for _, row in df.iterrows():
        wms_sku = clean_str(row.iloc[1]) if len(row) > 1 else ""
        product_name = clean_str(row.iloc[2]) if len(row) > 2 else ""
        lot = clean_str(row.iloc[3]) if len(row) > 3 else "NO_LOT"
        exp_date_str = clean_str(row.iloc[4]) if len(row) > 4 else ""
        barcode = clean_str(row.iloc[6]) if len(row) > 6 else ""
        warehouse = clean_str(row.iloc[7]) if len(row) > 7 else ""
        avail_qty = parse_number(row.iloc[10]) if len(row) > 10 else 0

        if not wms_sku and not barcode and not product_name:
            continue
        if avail_qty <= 0:
            continue

        exp_date = pd.to_datetime(exp_date_str, errors="coerce")

        parsed_rows.append({
            "WMS상품코드": wms_sku,
            "WMS상품명": product_name,
            "LOT": lot,
            "유통기한": exp_date,
            "바코드": barcode,
            "창고": warehouse,
            "가용재고": avail_qty,
        })

    return pd.DataFrame(parsed_rows)


def parse_oliveyoung_delivery(df_raw):
    """
    올리브영 납품확인서 목록 파싱
    """
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = " ".join(row.dropna().astype(str))
        if "상품코드" in row_str and "발주수량" in row_str:
            header_idx = idx
            break

    df = df_raw.copy()
    df.columns = df.iloc[header_idx].map(clean_str)
    df = df.iloc[header_idx + 1:].reset_index(drop=True)

    df = df[df["상품코드"].notna() & (df["상품코드"] != "")].copy()
    df["상품코드"] = df["상품코드"].apply(clean_str)
    
    req_col = [c for c in df.columns if "발주수량" in c]
    box_col = [c for c in df.columns if "BOX" in c and "입수" in c]
    
    df["발주수량(EA)"] = df[req_col[0]].apply(parse_number) if req_col else 0
    df["BOX입수"] = df[box_col[0]].apply(parse_number) if box_col else 1
    df["입고예정일"] = pd.to_datetime(df["입고예정일"], errors="coerce")

    return df

# ============================================================
# AUTO MATCHING & ALLOCATION LOGIC
# ============================================================
def allocate_inventory(delivery_df, wms_df, min_days=MIN_SHELF_LIFE_DAYS):
    """
    FEFO 기반 LOT 자동 매핑 및 유통기한/BOX입수/LOT쪼개짐 검증
    """
    wms_work = wms_df.copy()
    wms_work["가용재고_남은수량"] = wms_work["가용재고"]

    allocated_results = []

    for idx, d_row in delivery_df.iterrows():
        barcode = d_row.get("상품코드", "")
        req_qty = d_row.get("발주수량(EA)", 0)
        box_in = d_row.get("BOX입수", 1)
        target_date = d_row.get("입고예정일", pd.Timestamp.now())

        if pd.isna(target_date):
            target_date = pd.Timestamp.now()

        # 1차 바코드 매칭 -> 2차 SKU 매칭
        matched_wms = wms_work[(wms_work["바코드"] == barcode) | (wms_work["WMS상품코드"] == barcode)].copy()

        if not matched_wms.empty:
            matched_wms["잔여일수"] = (matched_wms["유통기한"] - target_date).dt.days
            # 유통기한 빠른 순(FEFO) 정렬
            valid_wms = matched_wms[matched_wms["잔여일수"] >= min_days].sort_values("유통기한")
            invalid_wms = matched_wms[matched_wms["잔여일수"] < min_days]
        else:
            valid_wms = pd.DataFrame()
            invalid_wms = pd.DataFrame()

        remaining_to_pick = req_qty
        picked_lots = []
        status = "NORMAL"

        if valid_wms.empty:
            if not invalid_wms.empty:
                status = "INVALID_SHELF_LIFE"
            else:
                status = "NO_STOCK"
        else:
            for w_idx, w_row in valid_wms.iterrows():
                if remaining_to_pick <= 0:
                    break
                
                avail = w_row["가용재고_남은수량"]
                if avail <= 0:
                    continue

                take = min(avail, remaining_to_pick)
                wms_work.loc[w_idx, "가용재고_남은수량"] -= take
                remaining_to_pick -= take

                exp_str = w_row["유통기한"].strftime("%Y-%m-%d") if pd.notna(w_row["유통기한"]) else "N/A"
                picked_lots.append({
                    "LOT": w_row["LOT"],
                    "유통기한": exp_str,
                    "수량": take,
                    "잔여일수": w_row["잔여일수"]
                })

            if remaining_to_pick > 0:
                status = "SHORTAGE"

        lot_summary = []
        exp_summary = []
        is_split = len(picked_lots) > 1

        for p in picked_lots:
            lot_summary.append(f"{p['LOT']}({int(p['수량'])}개)")
            exp_summary.append(p["유통기한"])

        status_flag = "🟢 정상출고"
        if status == "NO_STOCK":
            status_flag = "🔴 재고없음"
        elif status == "INVALID_SHELF_LIFE":
            status_flag = "⛔ [출고불가] 유통기한 1년6개월 미만"
        elif status == "SHORTAGE":
            status_flag = f"🟡 수량부족 (부족량: {int(remaining_to_pick)}EA)"
        elif is_split:
            status_flag = "⚠️ [주의] LOT 쪼개짐 (다중 LOT)"

        box_check = "OK"
        if box_in > 0 and (req_qty % box_in != 0):
            box_check = f"❌ 박스미달 (입수:{int(box_in)})"

        allocated_results.append({
            "순번": d_row.get("순번", idx + 1),
            "입고예정일": target_date.strftime("%Y-%m-%d") if pd.notna(target_date) else "",
            "센터": d_row.get("센터", ""),
            "바코드/상품코드": barcode,
            "상품명": d_row.get("상품명", ""),
            "발주수량(EA)": req_qty,
            "BOX입수": box_in,
            "매핑상태": status_flag,
            "박스입수체크": box_check,
            "매핑된 LOT": " / ".join(lot_summary) if lot_summary else "-",
            "유통기한": " / ".join(set(exp_summary)) if exp_summary else "-",
            "LOT split 여부": "O" if is_split else "X",
        })

    return pd.DataFrame(allocated_results), wms_work

# ============================================================
# STREAMLIT UI
# ============================================================
st.title("📦 올리브영 자동 출고 & LOT 매핑 시스템 (그레이스 3PL)")
st.caption("그레이스 WMS 재고현황과 올리브영 납품확인서를 자동 분석하여 유통기한 검증 및 LOT 쪼개짐을 체크합니다.")

st.sidebar.header("⚙️ 설정 옵션")
min_months = st.sidebar.slider("올리브영 납품 가능 최소 유통기한 (개월)", 6, 24, 18, 1)
min_days_limit = int(min_months * 30.4375)

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 그레이스 WMS 재고현황")
    # type 파라미터를 완전히 제거하여 .xls 확장자 업로드 차단을 해제함
    wms_file = st.file_uploader("WMS 재고현황 파일 (.xls, .xlsx 파일 그대로 업로드)", key="wms")

with col2:
    st.subheader("2. 올리브영 납품확인서")
    delivery_file = st.file_uploader("납품확인서 목록 파일 (.xlsx)", key="delivery")

if wms_file and delivery_file:
    try:
        wms_raw = load_uploaded_file(wms_file)
        delivery_raw = load_uploaded_file(delivery_file)

        wms_df = parse_grace_wms(wms_raw)
        delivery_df = parse_oliveyoung_delivery(delivery_raw)

        st.success(f"✅ 파일 자동 읽기 성공! (WMS 가용재고: {len(wms_df)}건 / 올리브영 납품항목: {len(delivery_df)}건)")

        if st.button("🚀 LOT 자동 매핑 및 출고 검증 실행", type="primary"):
            result_df, updated_wms = allocate_inventory(delivery_df, wms_df, min_days=min_days_limit)

            st.subheader("📊 자동 매핑 및 출고 검증 결과")

            k1, k2, k3, k4 = st.columns(4)
            normal_cnt = sum(result_df["매핑상태"].str.contains("정상출고"))
            split_cnt = sum(result_df["LOT split 여부"] == "O")
            invalid_cnt = sum(result_df["매핑상태"].str.contains("출고불가"))
            shortage_cnt = sum(result_df["매핑상태"].str.contains("부족|재고없음"))

            k1.metric("🟢 정상 출고 가능", f"{normal_cnt} 건")
            k2.metric("⚠️ LOT 쪼개짐 항목", f"{split_cnt} 건")
            k3.metric("⛔ 유통기한 부적합", f"{invalid_cnt} 건")
            k4.metric("🔴 재고 부족/없음", f"{shortage_cnt} 건")

            status_filter = st.multiselect("상태별 필터링", result_df["매핑상태"].unique(), default=result_df["매핑상태"].unique())
            filtered_result = result_df[result_df["매핑상태"].isin(status_filter)]

            st.dataframe(filtered_result, use_container_width=True, height=450)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                filtered_result.to_excel(writer, index=False, sheet_name="올리브영_LOT매핑결과")
                wms_df.to_excel(writer, index=False, sheet_name="WMS_원본재고")
            excel_data = output.getvalue()

            st.download_button(
                label="📥 매핑 결과 엑셀 다운로드",
                data=excel_data,
                file_name=f"올리브영_LOT매핑결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
else:
    st.info("💡 팀원들은 WMS에서 다운로드한 .xls 원본 파일을 수정 없이 그대로 올려주시면 됩니다.")
