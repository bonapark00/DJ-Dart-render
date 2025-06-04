import dart_fss as dart
import pandas as pd
from pandas import to_datetime

dart.set_api_key(api_key='8b1e1ecff1d195b34f0af2b7cc263e09275bfedf')

def extract_df(reports, separate = False): # 연결(False, 기본)

    # 결과 누적 리스트
    df_list = []

    col_label_ko_00 = ('[D431410] 단일 포괄손익계산서, 기능별 분류, 세후 - 연결 | Statement of comprehensive income, by function of expense - Consolidated financial statements (Unit: KRW)', 'label_ko')
    col_label_ko_01 = ('[D310000] 손익계산서, 기능별 분류 - 연결 | Income statement, by function of expense - Consolidated financial statements (Unit: KRW)', 'label_ko')
    col_label_ko_10 = ('[D431410] Statement of comprehensive income, by function of expense - Consolidated financial statements (Unit: KRW)', 'label_ko')
    col_label_ko_11 = ('[D310000] Income statement, by function of expense - Consolidated financial statements (Unit: KRW)', 'label_ko')


    prev_report_nm = ''

    for report in reports:
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

            cf_list = xbrl.get_income_statement(separate = separate)
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
                col[1][0] in ('연결재무제표', '별도재무제표') or \
                col == 'label_ko_0' or\
                col == 'label_ko_1']
            
            df = df[filter_columns]
            
            
            # 연도 추출 (보고서 제출일 기준)
            year = report.rcept_dt[:4]
            
            # 연도 정보를 컬럼으로 추가
            # df['year'] = year

            df_list.append(df)

        except Exception as e:
            print(f"❌ 오류 발생: {report.rcept_no} / {e}")
            continue

    # 모든 연도별 현금흐름표를 하나의 DF로 병합
    df_all = pd.concat(df_list, ignore_index=True)
    if 'label_ko_1' in df_all.columns:
        df_all['label_ko'] = df_all['label_ko_0'].combine_first(df_all['label_ko_1'])
        df_all.drop(columns = ['label_ko_0', 'label_ko_1'], inplace=True)
    else:
        df_all['label_ko'] = df_all['label_ko_0']
        df_all.drop(columns = ['label_ko_0'], inplace=True)
        
    # label_ko_0,1 정보들도 합쳐주기
    df_all = df_all.groupby('label_ko', as_index=False).first()


    # # 분기 조건을 만족하는 column만 남기기
    # quarter_columns = [
    #     col for col in df_all.columns
    #     if col != 'label_ko' and isinstance(col, tuple) and
    #     (to_datetime(col[0].split('-')[1]) - to_datetime(col[0].split('-')[0])).days <= 100
    # ]
    # print()

    # # label_ko는 유지하고, 필터링된 column만 선택
    # df_all = df_all[['label_ko'] + quarter_columns]

    cols = df_all.columns

    # 'label_ko'는 string이고 나머지는 tuple로 되어 있으므로 따로 분리
    label_col = [col for col in cols if col == 'label_ko']

    # 연결재무제표와 별도재무제표로 분리
    consol_cols = [col for col in cols if isinstance(col, tuple) and col[1][0] == '연결재무제표']
    separate_cols = [col for col in cols if isinstance(col, tuple) and col[1][0] == '별도재무제표']

    # 정렬: 최근 날짜 순 (내림차순)
    consol_cols_sorted = sorted(consol_cols, key=lambda x: x[0], reverse=True)
    separate_cols_sorted = sorted(separate_cols, key=lambda x: x[0], reverse=True)

    # 최종 컬럼 순서: label_ko + 정렬된 데이터 컬럼
    consol_final_cols = label_col + consol_cols_sorted
    separate_final_cols = label_col + separate_cols_sorted

    # 분리된 DataFrame
    df_consol = df_all[consol_final_cols]
    df_separate = df_all[separate_final_cols]

    return df_separate, df_consol

def df_merge(df_a001, df_a002, df_a003):


    # label_ko 컬럼 추출
    label_col = [col for col in df_a003.columns if col == 'label_ko'][0]

    # 기준 label_ko로 병합을 위한 기준 DF 생성
    df_base = df_a003[[label_col]].copy()

    # 기준 날짜
    endYear = df_a003.columns[1][0][:4]
    begYear = df_a003.columns[-1][0][:4]
    q1_cols, q2_cols, q3_cols, q4_cols = [], [], [], []

    for year in range(int(endYear), int(begYear)-1, -1):
        year = str(year)
    
        # separate = '별도재무제표'
        # connect = '연결재무제표'
        type_sc = df_a001.columns[1][1][0]

        # 날짜 포맷 기준 (시작일-종료일, 재무제표 구분)
        q1_col = (f'{year}0101-{year}0331', (type_sc,))
        q2_col = (f'{year}0401-{year}0630', (type_sc,))  # 계산 필요
        q3_col = (f'{year}0701-{year}0930', (type_sc,))
        q4_col = (f'{year}1001-{year}1231', (type_sc,))  # 계산 필요
        q123_col = (f'{year}0101-{year}0930', (type_sc,))
        y1_col = (f'{year}0101-{year}1231', (type_sc,))


        # Q1
        if q1_col in df_a003.columns:
            df_base[f'{year}_Q1'] = df_a003.get(q1_col)
            q1_cols.append(f'{year}_Q1')
        # Q2
        if q2_col in df_a002.columns:
            df_base[f'{year}_Q2'] = df_a002.get(q2_col)
            q2_cols.append(f'{year}_Q2')

        # Q3
        if q3_col in df_a003.columns:
            df_base[f'{year}_Q3'] = df_a003.get(q3_col)
            q3_cols.append(f'{year}_Q3')

        # Q4
        if y1_col in df_a001.columns and q123_col in df_a003.columns:
            df_base[f'{year}_Q4'] = df_a001.get(y1_col) - df_a003.get(q123_col)
            q4_cols.append(f'{year}_Q4')


        cols = df_base.columns.tolist()

        # 'label_ko'는 따로 빼고 나머지 분기 컬럼만 추출
        quarter_cols = [col for col in cols if col != 'label_ko']

        # 분기 컬럼을 날짜 기준으로 정렬 (최근 순서)
        sorted_quarters = sorted(
            quarter_cols,
            key=lambda x: (int(x[:4]), int(x[-1])),  # ('2024_Q3' → (2024, 3))
            reverse=True
        )

        # 새 컬럼 순서: label_ko + 최근 분기 순서
        new_cols = ['label_ko'] + sorted_quarters

        # 컬럼 순서 재정렬
        df_base = df_base[new_cols]
    return df_base, df_base[['label_ko']+q1_cols], df_base[['label_ko']+q2_cols], df_base[['label_ko']+q3_cols], df_base[['label_ko']+q4_cols], 
    
def get_income_by_name(corp_name, corp_market, bgn_de, end_de):
        
    # 삼성전자 code
    # corp_code = '00336570'
    # corp_name = '원텍'
    # corp_market = 'K' # 'Y': 코스피, 'K': 코스닥, 'N': 코넥스, 'E': 기타


    # 모든 상장된 기업 리스트 불러오기
    corp_list = dart.get_corp_list() 
    clists = corp_list.find_by_corp_name(corp_name=corp_name, exactly=True, market = corp_market)
    print(clists)
    corp = clists[0]

    # 반기보고서 검색
    reports_a001 = corp.search_filings(bgn_de=bgn_de, end_de=end_de, pblntf_detail_ty='a001')
    reports_a002 = corp.search_filings(bgn_de=bgn_de, end_de=end_de, pblntf_detail_ty='a002')
    reports_a003 = corp.search_filings(bgn_de=bgn_de, end_de=end_de, pblntf_detail_ty='a003')



    # a001
    df_a001_sep, df_a001_con = extract_df(reports_a001)
    # a002
    df_a002_sep, df_a002_con = extract_df(reports_a002)
    # a003
    df_a003_sep, df_a003_con = extract_df(reports_a003)

    df_con_total, df_con_q1, df_con_q2, df_con_q3, df_con_q4 = df_merge(df_a001_con, df_a002_con, df_a003_con)
    df_sep_total, df_sep_q1, df_sep_q2, df_sep_q3, df_sep_q4, = df_merge(df_a001_sep, df_a002_sep, df_a003_sep)
    
    
    dfs = [
        ('전체분기_별도', df_sep_total),
        ('전체분기_연결', df_con_total),
        ('Q1_별도', df_sep_q1),
        ('Q2_별도', df_sep_q2),
        ('Q3_별도', df_sep_q3),
        ('Q4_별도', df_sep_q4),
        ('Q1_연결', df_con_q1),
        ('Q2_연결', df_con_q2),
        ('Q3_연결', df_con_q3),
        ('Q4_연결', df_con_q4),
    ]

    return dfs



    from openpyxl.utils import get_column_letter
    from openpyxl.styles import numbers
    from openpyxl import load_workbook

    # Excel 파일로 저장
    with pd.ExcelWriter('포과손익계산서.xlsx', engine='openpyxl') as writer:
        df_sep_total.to_excel(writer, sheet_name='전체분기_별도', index=False)
        df_con_total.to_excel(writer, sheet_name='전체분기_연결', index=False)
        df_sep_q1.to_excel(writer, sheet_name='Q1_별도', index=False)
        df_sep_q2.to_excel(writer, sheet_name='Q2_별도', index=False)
        df_sep_q3.to_excel(writer, sheet_name='Q3_별도', index=False)
        df_sep_q4.to_excel(writer, sheet_name='Q4_별도', index=False)
        df_con_q1.to_excel(writer, sheet_name='Q1_연결', index=False)
        df_con_q2.to_excel(writer, sheet_name='Q2_연결', index=False)
        df_con_q3.to_excel(writer, sheet_name='Q3_연결', index=False)
        df_con_q4.to_excel(writer, sheet_name='Q4_연결', index=False)

    # 저장한 파일을 다시 불러와 포맷 적용
    wb = load_workbook('포과손익계산서.xlsx')

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        
        # A열 (label_ko) 너비 넓히기
        ws.column_dimensions['A'].width = 40
        
        
        # B열부터 마지막 열까지 숫자 형식 적용 (천 단위 쉼표)
        for col in range(2, ws.max_column + 1):
            col_letter = get_column_letter(col)
            for row in range(2, ws.max_row + 1):  # header 제외
                cell = ws[f'{col_letter}{row}']
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0'  # 천 단위 쉼표
            ws.column_dimensions[col_letter].width = 20


    # 다시 저장
    wb.save('포과손익계산서.xlsx')