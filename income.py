import dart_fss as dart
import pandas as pd
import gc

# DART API 키 설정 (본인의 API 키로 변경하세요)
# dart.set_api_key(api_key='YOUR_DART_API_KEY')
# 보안을 위해 실제 서비스에서는 환경 변수 등으로 관리하는 것이 좋습니다.
try:
    dart.set_api_key(api_key='8b1e1ecff1d195b34f0af2b7cc263e09275bfedf')
except Exception as e:
    print(f"DART API 키 설정 오류: {e}. API 키를 확인해주세요.")

def extract_df_for_year(reports, year, separate=False):
    """
    주어진 보고서 리스트에서 특정 연도의 재무제표 데이터를 추출합니다.
    """
    df_list = []
    
    # 컬럼 이름을 위한 정의 (DART API 응답 형식에 따라 변경될 수 있음)
    col_label_ko_00 = ('[D431410] 단일 포괄손익계산서, 기능별 분류, 세후 - 연결 | Statement of comprehensive income, by function of expense - Consolidated financial statements (Unit: KRW)', 'label_ko')
    col_label_ko_01 = ('[D310000] 손익계산서, 기능별 분류 - 연결 | Income statement, by function of expense - Consolidated financial statements (Unit: KRW)', 'label_ko')
    col_label_ko_10 = ('[D431410] Statement of comprehensive income, by function of expense - Consolidated financial statements (Unit: KRW)', 'label_ko')
    col_label_ko_11 = ('[D310000] Income statement, by function of expense - Consolidated financial statements (Unit: KRW)', 'label_ko')

    prev_report_nm = ''

    # 해당 연도에 해당하는 보고서만 필터링
    # DART API에서 직접 연도별 필터링이 가능하면 더욱 효율적이지만,
    # 여기서는 이미 가져온 reports 리스트에서 연도 필터링을 수행합니다.
    # report.rcept_dt (접수일자)를 기준으로 해당 연도의 보고서만 처리
    reports_for_year = [r for r in reports if r.rcept_dt.startswith(str(year))]

    for report in reports_for_year:
        try:
            # 개정정정이 존재하는 경우, 그 다음 기존 보고서는 건너뛴다(가장 최근 데이터만 사용)
            if prev_report_nm == report.report_nm[-9:]:
                continue
            else:
                prev_report_nm = report.report_nm[-9:]

            xbrl = report.xbrl

            # 연결재무제표 있는 경우만 수집
            if not xbrl.exist_consolidated():
                continue

            cf_list = xbrl.get_income_statement(separate=separate)
            if not cf_list:
                continue

            cf = cf_list[0]
            df = cf.to_DataFrame(show_class=False)

            new_columns = []
            for col in df.columns:
                if col == col_label_ko_00 or col == col_label_ko_01:
                    new_columns.append('label_ko_0')
                elif col == col_label_ko_10 or col == col_label_ko_11:
                    new_columns.append('label_ko_1')
                else:
                    new_columns.append(col)
            df.columns = new_columns

            filter_columns = [col for col in df.columns if \
                              (isinstance(col, tuple) and (col[1][0] in ('연결재무제표', '별도재무제표'))) or \
                              col == 'label_ko_0' or \
                              col == 'label_ko_1']
            
            # 현재 처리 중인 연도와 관련된 컬럼만 필터링합니다.
            # 예를 들어, ('20240101-20240331', ('연결재무제표',)) 와 같은 컬럼
            current_year_cols = [col for col in filter_columns if isinstance(col, tuple) and col[0].startswith(str(year))]
            # label_ko_0 또는 label_ko_1 컬럼도 포함
            current_year_cols += [col for col in filter_columns if not isinstance(col, tuple)]

            df = df[current_year_cols]

            df_list.append(df)

        except Exception as e:
            print(f"❌ 오류 발생 (보고서 {report.rcept_no}, 연도 {year}): {e}")
            continue

    if not df_list:
        return pd.DataFrame(), pd.DataFrame() # 빈 DataFrame 반환

    # 모든 연도별 현금흐름표를 하나의 DF로 병합 (이 시점에서는 해당 연도 데이터만 있음)
    df_all = pd.concat(df_list, ignore_index=True)
    
    del df_list
    gc.collect()

    if 'label_ko_1' in df_all.columns:
        df_all['label_ko'] = df_all['label_ko_0'].combine_first(df_all['label_ko_1'])
        df_all.drop(columns=['label_ko_0', 'label_ko_1'], inplace=True)
    else:
        df_all['label_ko'] = df_all['label_ko_0']
        df_all.drop(columns=['label_ko_0'], inplace=True)

    df_all = df_all.groupby('label_ko', as_index=False).first()

    cols = df_all.columns

    label_col = [col for col in cols if col == 'label_ko']

    # 연결재무제표와 별도재무제표로 분리 (해당 연도 데이터만 포함)
    consol_cols = [col for col in cols if isinstance(col, tuple) and col[1][0] == '연결재무제표']
    separate_cols = [col for col in cols if isinstance(col, tuple) and col[1][0] == '별도재무제표']

    consol_cols_sorted = sorted(consol_cols, key=lambda x: x[0], reverse=True)
    separate_cols_sorted = sorted(separate_cols, key=lambda x: x[0], reverse=True)

    consol_final_cols = label_col + consol_cols_sorted
    separate_final_cols = label_col + separate_cols_sorted

    df_consol = df_all[consol_final_cols]
    df_separate = df_all[separate_final_cols]
    
    del df_all
    gc.collect()

    return df_separate, df_consol


def df_merge_for_year(df_a001, df_a002, df_a003, year):
    """
    특정 연도의 분기별 데이터를 병합합니다.
    """
    label_col = [col for col in df_a003.columns if col == 'label_ko'][0]
    df_base = df_a003[[label_col]].copy()

    str_year = str(year)
    q1_cols, q2_cols, q3_cols, q4_cols = [], [], [], []

    # df_a001, df_a002, df_a003에 해당 연도의 데이터만 있다고 가정합니다.
    
    # 재무제표 구분 (연결/별도) - 컬럼 구조에 따라 동적으로 가져옵니다.
    # 만약 df_a001, df_a002, df_a003이 비어있을 수 있으므로 체크
    type_sc = None
    if not df_a001.empty and len(df_a001.columns) > 1 and isinstance(df_a001.columns[1], tuple):
        type_sc = df_a001.columns[1][1][0]
    elif not df_a002.empty and len(df_a002.columns) > 1 and isinstance(df_a002.columns[1], tuple):
        type_sc = df_a002.columns[1][1][0]
    elif not df_a003.empty and len(df_a003.columns) > 1 and isinstance(df_a003.columns[1], tuple):
        type_sc = df_a003.columns[1][1][0]

    if type_sc is None: # 재무제표 타입 정보를 찾을 수 없는 경우
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    q1_col = (f'{str_year}0101-{str_year}0331', (type_sc,))
    q2_col = (f'{str_year}0401-{str_year}0630', (type_sc,))
    q3_col = (f'{str_year}0701-{str_year}0930', (type_sc,))
    q4_col_full_year = (f'{str_year}0101-{str_year}1231', (type_sc,)) # 연간
    q4_col_q123 = (f'{str_year}0101-{str_year}0930', (type_sc,)) # 1-3분기 누적

    # Q1
    if q1_col in df_a003.columns:
        df_base[f'{str_year}_Q1'] = df_a003.get(q1_col)
        q1_cols.append(f'{str_year}_Q1')

    # Q2
    if q2_col in df_a002.columns:
        df_base[f'{str_year}_Q2'] = df_a002.get(q2_col)
        q2_cols.append(f'{str_year}_Q2')

    # Q3
    if q3_col in df_a003.columns:
        df_base[f'{str_year}_Q3'] = df_a003.get(q3_col)
        q3_cols.append(f'{str_year}_Q3')

    # Q4 (연간 데이터 - 1-3분기 누적 데이터)
    if q4_col_full_year in df_a001.columns and q4_col_q123 in df_a003.columns:
        df_base[f'{str_year}_Q4'] = df_a001.get(q4_col_full_year) - df_a003.get(q4_col_q123)
        q4_cols.append(f'{str_year}_Q4')

    cols = df_base.columns.tolist()
    quarter_cols = [col for col in cols if col != 'label_ko']

    if quarter_cols: # 분기 데이터가 있는 경우만 정렬
        sorted_quarters = sorted(
            quarter_cols,
            key=lambda x: (int(x[:4]), int(x[-1])),
            reverse=True
        )
        new_cols = ['label_ko'] + sorted_quarters
        df_base = df_base[new_cols]
    else: # 분기 데이터가 없는 경우 label_ko만 남김
        df_base = df_base[['label_ko']]
    
    # 각 분기별 DF를 생성 (해당 연도 데이터만 포함)
    df_total = df_base
    df_q1 = df_base[['label_ko'] + q1_cols] if q1_cols else pd.DataFrame(columns=['label_ko'])
    df_q2 = df_base[['label_ko'] + q2_cols] if q2_cols else pd.DataFrame(columns=['label_ko'])
    df_q3 = df_base[['label_ko'] + q3_cols] if q3_cols else pd.DataFrame(columns=['label_ko'])
    df_q4 = df_base[['label_ko'] + q4_cols] if q4_cols else pd.DataFrame(columns=['label_ko'])

    return df_total, df_q1, df_q2, df_q3, df_q4


def get_income_by_name(corp_name, corp_market, bgn_de, end_de, filepath):
    """
    기업의 재무제표를 연도별로 조회하여 엑셀 파일에 저장합니다.
    """
    corp_list = dart.get_corp_list()
    clists = corp_list.find_by_corp_name(corp_name=corp_name, exactly=True, market=corp_market)
    if not clists:
        raise ValueError(f"기업명 '{corp_name}' ({corp_market})을(를) 찾을 수 없습니다.")
    corp = clists[0]

    # 전체 기간에 대한 보고서 리스트를 한 번에 가져옵니다.
    # 각 보고서 객체 자체는 가볍지만, XBRL 데이터를 로드할 때 메모리를 많이 사용합니다.
    # 따라서 extract_df_for_year에서 필요한 연도의 XBRL만 로드하도록 합니다.
    reports_a001_all = corp.search_filings(bgn_de=bgn_de, end_de=end_de, pblntf_detail_ty='a001')
    reports_a002_all = corp.search_filings(bgn_de=bgn_de, end_de=end_de, pblntf_detail_ty='a002')
    reports_a003_all = corp.search_filings(bgn_de=bgn_de, end_de=end_de, pblntf_detail_ty='a003')

    start_year = int(bgn_de[:4])
    end_year = int(end_de[:4])
    
    # ExcelWriter는 파일 핸들을 유지하며 시트를 추가하므로, 전체 과정에서 한 번만 생성합니다.
    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        for year in range(end_year, start_year - 1, -1): # 최신 연도부터 과거 연도 순으로 처리
            print(f"📡 {year}년도 데이터 처리 중...")
            
            # 각 연도별로 데이터를 추출
            # extract_df_for_year는 해당 연도의 보고서만 필터링하여 처리합니다.
            df_a001_sep, df_a001_con = extract_df_for_year(reports_a001_all, year, separate=False)
            df_a002_sep, df_a002_con = extract_df_for_year(reports_a002_all, year, separate=False)
            df_a003_sep, df_a003_con = extract_df_for_year(reports_a003_all, year, separate=False)
            
            # 각 연도별로 병합된 데이터 프레임 생성
            df_con_total, df_con_q1, df_con_q2, df_con_q3, df_con_q4 = df_merge_for_year(df_a001_con, df_a002_con, df_a003_con, year)
            df_sep_total, df_sep_q1, df_sep_q2, df_sep_q3, df_sep_q4 = df_merge_for_year(df_a001_sep, df_a002_sep, df_a003_sep, year)

            # 데이터가 비어있지 않은 경우에만 엑셀에 저장
            if not df_sep_total.empty:
                df_sep_total.to_excel(writer, sheet_name=f'별도_전체_{year}'[:31], index=False)
            if not df_con_total.empty:
                df_con_total.to_excel(writer, sheet_name=f'연결_전체_{year}'[:31], index=False)
            if not df_sep_q1.empty:
                df_sep_q1.to_excel(writer, sheet_name=f'별도_Q1_{year}'[:31], index=False)
            if not df_sep_q2.empty:
                df_sep_q2.to_excel(writer, sheet_name=f'별도_Q2_{year}'[:31], index=False)
            if not df_sep_q3.empty:
                df_sep_q3.to_excel(writer, sheet_name=f'별도_Q3_{year}'[:31], index=False)
            if not df_sep_q4.empty:
                df_sep_q4.to_excel(writer, sheet_name=f'별도_Q4_{year}'[:31], index=False)
            if not df_con_q1.empty:
                df_con_q1.to_excel(writer, sheet_name=f'연결_Q1_{year}'[:31], index=False)
            if not df_con_q2.empty:
                df_con_q2.to_excel(writer, sheet_name=f'연결_Q2_{year}'[:31], index=False)
            if not df_con_q3.empty:
                df_con_q3.to_excel(writer, sheet_name=f'연결_Q3_{year}'[:31], index=False)
            if not df_con_q4.empty:
                df_con_q4.to_excel(writer, sheet_name=f'연결_Q4_{year}'[:31], index=False)

            # 현재 연도와 관련된 모든 DataFrame을 명시적으로 삭제
            del df_a001_sep, df_a001_con, df_a002_sep, df_a002_con, df_a003_sep, df_a003_con
            del df_con_total, df_con_q1, df_con_q2, df_con_q3, df_con_q4
            del df_sep_total, df_sep_q1, df_sep_q2, df_sep_q3, df_sep_q4
            gc.collect() # 가비지 컬렉션 실행

    # 파일 경로만 반환 (app.py에서 다운로드 및 후처리)
    return filepath