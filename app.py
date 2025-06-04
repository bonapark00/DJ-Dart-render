import streamlit as st
import pandas as pd
import tempfile
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ⚠️ Render 환경에서는 상대 import 기준
from income import get_income_by_name  # income.py가 같은 폴더에 있어야 함

st.set_page_config(page_title="DART 재무제표 추출기", layout="wide")
st.title("📊 DART 재무제표 추출기")

# 입력창
corp_name = st.text_input("기업명", "삼성전자")
corp_market = st.text_input("기업 구분 ('Y': 코스피, 'K': 코스닥, 'N': 코넥스, 'E': 기타)", "Y")
bgn_de = st.text_input("시작일 (YYYYMMDD)", "20220101")
end_de = st.text_input("종료일 (YYYYMMDD)", "20241231")

# 버튼
if st.button("📥 엑셀 파일 생성 및 다운로드"):

    with st.spinner("📡 데이터를 조회하고 있습니다..."):
        try:
            # 재무제표 가져오기
            dfs = get_income_by_name(
                corp_name=corp_name,
                corp_market=corp_market,
                bgn_de=bgn_de,
                end_de=end_de
            )

            # 임시 엑셀 파일 생성
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                filepath = tmp.name

            # 엑셀 저장
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                for sheet_name, df in dfs:
                    df.to_excel(writer, sheet_name=sheet_name[:31], index=False)  # 시트 이름 최대 31자

            # 서식 적용
            wb = load_workbook(filepath)
            for sheetname in wb.sheetnames:
                ws = wb[sheetname]
                ws.column_dimensions['A'].width = 40  # label_ko 열
                for col in range(2, ws.max_column + 1):
                    col_letter = get_column_letter(col)
                    max_width = 10
                    for row in range(2, ws.max_row + 1):
                        cell = ws.cell(row=row, column=col)
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = '#,##0'
                            formatted = f"{cell.value:,.0f}"
                            max_width = max(max_width, len(formatted) + 2)
                    ws.column_dimensions[col_letter].width = max_width
            wb.save(filepath)

            # 다운로드 버튼
            with open(filepath, 'rb') as f:
                st.success("✅ 엑셀 파일이 준비되었습니다!")
                st.download_button(
                    label="📥 엑셀 다운로드",
                    data=f,
                    file_name=f"{corp_name}_포괄손익계산서.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"❌ 오류 발생: {e}")