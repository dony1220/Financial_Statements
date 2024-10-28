import pandas as pd
import numpy as np
from dash import Dash, dcc, html, Input, Output, dash_table, html
import FinanceDataReader as fdr
import re
import base64
import unicodedata
import datetime
import csv
import json


# 데이터를 읽어옴
financial_data_path = '통합재무상태표.csv'  # 재무상태표 파일 경로
income_statement_data_path = '통합손익계산서.csv'  # 손익계산서 파일 경로
cashflow_data_path = '통합현금흐름표.csv'  # 현금흐름표 파일 경로

financial_data = pd.read_csv(financial_data_path, low_memory=False)
income_statement_data = pd.read_csv(income_statement_data_path, low_memory=False)
cashflow_data = pd.read_csv(cashflow_data_path, low_memory=False)

# 하이픈과 NaN 값을 대체
financial_data.replace('-', np.nan, inplace=True)
income_statement_data.replace('-', np.nan, inplace=True)
cashflow_data.replace('-', np.nan, inplace=True)

# 금액 컬럼들이 문자열로 되어 있으므로 이를 숫자로 변환하고 백만 원 단위로 변환
numeric_columns_financial = ['당기', '당기 1분기말', '당기 반기말', '당기 3분기말']
numeric_columns_income = ['당기', '당기 1분기 3개월', '당기 반기 3개월', '당기 3분기 3개월']
numeric_columns_cashflow = ['당기', '당기 1분기', '당기 반기', '당기 3분기']

def convert_to_numeric(value):
    if isinstance(value, str):
        value = re.sub(r'[(),]', '', value)  # 괄호와 콤마를 제거
        return pd.to_numeric(value, errors='coerce')  # 숫자로 변환
    return value

def normalize_item_name(item_name):
    if isinstance(item_name, str):
        item_name = item_name.replace(' ', '').replace('　', '')  # 모든 항목명에서 띄어쓰기를 제거
        if item_name == "당기손익-공정가치금융자산":
            return "당기손익-공정가치측정금융자산"
        elif item_name in ["비지배주주지분", 'II.비지배지분', '총포괄손익,비지배지분', ' II.비지배지분', '2. 비지배지분', '2. 비지배지분', '포괄손익, 비지배지분',
                           '         포괄손익, 비지배지분', '총 포괄손익, 비지배지분', '총 포괄손익, 비지배지분', '포괄손익, 비지배지분', '포괄손익,비지배지분',
                           '비지배지분순이익']:
            return "비지배지분"
        # if re.match(r'.*비지배지분', item_name):
        #     return '비지베지분'
        elif item_name == "이익잉여금(결손금)":
            return "이익잉여금"
        elif item_name in ["부채와자본총계", "부채및자본총계", '자본및부채총계']:
            return "자본과부채총계"
        elif item_name in ['지배기업소유지분', '지배기업의소유주에게귀속되는자본', '지배기업의소유주지분', 'I.지배기업의소유주에게귀속되는자본', '지배기업의소유주지본',
                           '지배기업소유주에게귀속되는자본', '지배기업의소유주', '지배기업의소유지분', 'I. 지배기업소유주지분', '1. 지배기업소유주지분',
                           '지배기업의 소유주지분', '   1. 지배기업소유주', '지배기업의 소유주에게 귀속되는 자본', '지배주주지분', '지배기업소유주',
                           '총 포괄손익, 지배기업의 소유주에게 귀속되는 지분', '포괄손익, 지배기업의 소유주에게 귀속되는 지분', '총 포괄손익, 지배기업의 소유주에게 귀속되는 지분',
                           '포괄손익, 지배기업의 소유주에게 귀속되는 지분', '포괄손익,지배기업의소유주에게귀속되는지분', '지배기업의소유주에게귀속되는자본', '총포괄손익,지배기업의 소유주에게 귀속되는 지분',
                           "총포괄손익,지배기업의소유주에게귀속되는지분", '총포괄손익,지배기업의소유주에게귀속되는지분', '지배기업지분', '지배기업소유주지분순이익']:
            return '지배기업소유주지분'
        elif item_name == '재고자산':
            return '유동재고자산'
        elif item_name in ['분기순이익(손실)', '분기순이익', '반기순이익', '반기순이익(손실)', 'XI.반기순이익(손실)', '당기순손익',
                           'XI.반기순이익', 'XI.분기순이익', '반기의순이익', 'XI.당기순이익', '분기의순이익', '반기순손익', 'XI.반기순이익(손실)',
                           '당기순이익(손실)', '연결반기순이익', '반기순손실', '당기순손실', '분기순손실', 'Ⅴ.당기순이익(손실)', '연결당기순이익(손실)', '당기의 순이익', '당기의순이익',
                           '당기순이익(손실)(A)', 'Ⅷ.당(분)기연결순이익', 'Ⅷ.당(반)기연결순이익']:
            return '당기순이익'
        elif item_name in ['당기총포괄손익', '분기총포괄손익', '총포괄손익', '총포괄손익(*3)', '총포괄손익']:
            return '당기총포괄손실'
        elif item_name in ['Ⅰ.유동자산', 'I. 유동자산', 'I. 유동자산']:
            return '유동자산'
        elif item_name in ['Ⅱ.비유동자산', 'II. 비유동자산', 'II. 비유동자산']:
            return '비유동자산'
        elif item_name in ['Ⅰ.유동부채', 'I. 유동부채', 'I. 유동부채']:
            return '유동부채'
        elif item_name in ['Ⅱ.비유동부채', 'II. 비유동부채', 'II. 비유동부채']:
            return '비유동부채'
        elif item_name in ['(1)자본금']:
            return '자본금'
        elif item_name in ['(2).자본잉여금', '(2)자본잉여금']:
            return '자본잉여금'
        elif item_name in ['(4)기타포괄손익누계액']:
            return '기타포괄손익누계액'
        elif item_name in ['(5)이익잉여금']:
            return '이익잉여금'
        elif item_name in ['반기말자본', '분기말자본', '당기말자본', '분기말', '반기말', '기말', '자본총계']:
            return '자본총계'
        elif item_name in ['XIII.총포괄이익', '총포괄손익합계', 'XⅢ.총포괄이익(손실)', 'XⅢ.총포괄손익', 'XI.총포괄손익',
                           '총포괄이익(손실)', '반기총포괄손익', '분기총포괄이익', '총포괄손익']:
            return '총포괄이익'
        elif item_name in ['Ⅴ.영업이익', '영업이익(손실)', '영업손실', 'V. 영업이익', 'V. 영업이익', 'Ⅲ.영업이익(손실)', '영업손익', '영업이익(손실)',
                           'Ⅲ.영업이익']:
            return '영업이익'
        elif item_name in ['Ⅳ.판매비와관리비', ' IV. 판매비와관리비']:
            return '판매비와관리비'
        elif item_name in ['비지배주주포괄이익(손실)']:
            return '비지배주주포괄이익'
        elif item_name in ['Ⅱ.매출원가']:
            return '매출원가'
        elif item_name in ['Ⅲ.매출총이익', '매출총이익(손실)']:
            return '매출총이익'
        elif item_name in ['Ⅹ.법인세비용']:
            return '법인세비용'
        elif item_name in ['XII.법인세비용차감후기타포괄손익']:
            return '법인세비용차감후기타포괄손익'
        # elif item_name in ['XI.반기순이익', 'XI.분기순이익', '반기의순이익', 'XI.당기순이익', '분기의순이익', '반기순손익', 'XI.반기순이익(손실)']:
        #     return '당기의순이익'
        elif item_name in ['수익(매출액)', 'Ⅰ.매출액', '매출', '영업수익', 'I. 매출액', 'I.매출액', '영업수익(매출액)', '매출및지분법손익', 'Ⅰ.영업수익', '이자수익(매출액']:
            return '매출액'
        elif item_name in ['부체총계']:
            return '부채총계'
        elif item_name in ['자산총계']:
            return '자산총계'
        elif item_name in ['기본주당이익(손실)', '기본주당이익', '보통주기본주당이익', '기본주당반기순이익(손실)', '기본주당분기순이익(손실)']:
            return 'EPS'
    return item_name

# 항목명 통일
financial_data['항목명'] = financial_data['항목명'].apply(normalize_item_name)
income_statement_data['항목명'] = income_statement_data['항목명'].apply(normalize_item_name)
cashflow_data['항목명'] = cashflow_data['항목명'].apply(normalize_item_name)

# 금액을 숫자로 변환하고 백만 원 단위로 변환
for col in numeric_columns_financial:
    financial_data[col] = financial_data[col].apply(convert_to_numeric).apply(lambda x: x / 100_000_000 if pd.notna(x) else np.nan)

for col in numeric_columns_income:
    income_statement_data[col] = income_statement_data[col].apply(convert_to_numeric).apply(lambda x: x / 100_000_000 if pd.notna(x) else np.nan)

for col in numeric_columns_cashflow:
    cashflow_data[col] = cashflow_data[col].apply(convert_to_numeric).apply(lambda x: x / 100_000_000 if pd.notna(x) else np.nan)

# 결산기준일에서 연도 추출
financial_data['결산연도'] = pd.to_datetime(financial_data['결산기준일'], errors='coerce').dt.year
income_statement_data['결산연도'] = pd.to_datetime(income_statement_data['결산기준일'], errors='coerce').dt.year
cashflow_data['결산연도'] = pd.to_datetime(cashflow_data['결산기준일'], errors='coerce').dt.year

# 데이터 정렬
# financial_data.sort_values(by=['회사명', '보고서종류'], inplace=True)
# income_statement_data.sort_values(by=['회사명', '보고서종류'], inplace=True)
# cashflow_data.sort_values(by=['회사명', '보고서종류'], inplace=True)

# 재무제표명(연결재무제표, 개별재무제표) 드롭다운 옵션 추출
statement_type_options = [{'label': statement, 'value': statement} for statement in financial_data['재무제표명'].unique()]

# 재무제표 종류 옵션 생성 (재무상태표, 손익계산서, 현금흐름표)


# 보고서종류별로 사용할 컬럼을 매핑하는 함수 (재무상태표, 손익계산서, 현금흐름표 분리)
def get_columns_by_report_type(report_type, statement_type):
    if statement_type == '재무상태표':
        if report_type == '사업보고서':
            return '당기'
        elif report_type == '3분기보고서':
            return '당기 3분기말'
        elif report_type == '반기보고서':
            return '당기 반기말'
        elif report_type == '1분기보고서':
            return '당기 1분기말'
    elif statement_type == '손익계산서':
        if report_type == '사업보고서':
            return '당기'
        elif report_type == '3분기보고서':
            return '당기 3분기 3개월'
        elif report_type == '반기보고서':
            return '당기 반기 3개월'
        elif report_type == '1분기보고서':
            return '당기 1분기 3개월'
    elif statement_type == '현금흐름표':
        if report_type == '사업보고서':
            return '당기'
        elif report_type == '3분기보고서':
            return '당기 3분기'
        elif report_type == '반기보고서':
            return '당기 반기'
        elif report_type == '1분기보고서':
            return '당기 1분기'
    return None

# 변화율 계산 함수
def calculate_change(current, previous):
    try:
        if pd.notna(current) and pd.notna(previous) and previous != 0:
            return (current - previous) / abs(previous) * 100
    except (TypeError, ValueError):
        return np.nan
    return np.nan

# 변화율에 따른 화살표와 괄호 추가, 색상 표시 함수
def format_change(change):
    if pd.notna(change):
        if change > 0:
            return f"({change:.1f}%) ▲", 'red'  # 상승: 빨간색
        elif change < 0:
            return f"({change:.1f}%) ▼", 'blue'  # 하락: 파란색
    return '-', 'black'  # 변화 없음

# Dash 앱 초기화
external_scripts = [
    {
        'src': 'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js'
    }
]
app = Dash(__name__, external_scripts=external_scripts)

# 회사명 옵션 생성
company_options = [{'label': company, 'value': company} for company in financial_data['회사명'].unique()]

# 보고서 종류 옵션 생성
report_options = [{'label': report, 'value': report} for report in financial_data['보고서종류'].unique()]

year_options = [{'label': str(year), 'value': year} for year in range(2020, 2025)]

financial_statement_options = [
    {'label': '재무상태표', 'value': '재무상태표'},
    {'label': '손익계산서', 'value': '손익계산서'},
    {'label': '현금흐름표', 'value': '현금흐름표'}
]


min_rate_change = 10
max_rate_change = 50
step_rate_change = 10




# 앱 레이아웃 정의
app.layout = html.Div([
    html.H1("2020 ~ 2024 재무 대시보드"),

    html.Hr(),  # 구분선 추가
    
    # 재무제표종류 선택 드롭다운
    html.Label("재무제표종류 선택"),
    dcc.Dropdown(id='financial-statement-dropdown', options=financial_statement_options, placeholder="재무제표종류를 선택하세요", style={'margin-bottom': '20px'}),

    # 회사명 선택 드롭다운
    html.Label("회사명 선택"),
    dcc.Dropdown(id='company-dropdown', options=company_options, placeholder="회사를 선택하세요", style={'margin-bottom': '20px'}),
    
    # 보고서 종류 선택 드롭다운
    html.Label("보고서종류 선택"),
    dcc.Dropdown(id='report-dropdown', options=report_options, placeholder="보고서를 선택하세요", style={'margin-bottom': '20px'}),

    # 재무제표명 선택 드롭다운
    html.Label("재무제표명 선택"),
    dcc.Dropdown(id='statement-type-dropdown', options=statement_type_options, placeholder="재무제표명을 선택하세요", style={'margin-bottom': '20px'}),

    # 연도 선택 드롭다운 추가
    html.Label("연도 선택(비교연도)"),
    dcc.Dropdown(id='year-dropdown', options=year_options, placeholder="연도를 선택하세요 (선택하지 않으면 전체 연도 표시)", style={'margin-bottom': '20px'}),


    html.Hr(),

    # 선택한 정보 및 테이블 표시 영역
    html.Div(id='table-and-info', children=[
        html.Div(id='selected-info', style={'text-align': 'center', 'font-size': '16px', 'margin-bottom': '20px', 'white-space': 'pre-line'}),
        html.Div(id='financial-data-table', style={'border': '1px solid black', 'padding': '10px', 'border-radius': '10px'}),
        html.Div("단위 : 억(원)", style={'text-align': 'right', 'font-size': '12px', 'margin-top': '5px', 'color': 'gray'})  # 단위 표시
    ], style={'border': '1px solid black', 'padding': '20px', 'border-radius': '10px'}),

    # CSV 다운로드 버튼 추가
    html.Button("Download Table as CSV", id="download-csv-btn", style={'margin-top': '20px'}),

    dcc.Download(id="download-dataframe-csv"),

    # 다운로드 버튼 추가 (PNG 저장)
    html.Button("Download Table as PNG", id="download-btn", style={'margin-top': '20px'}),

    html.Hr(),

    html.H1("등락률에 따른 회사 필터링"),
    html.Hr(),
    html.Label("재무제표 종류 선택"),
    dcc.Dropdown(id='independent-financial-statement-dropdown', options=financial_statement_options, placeholder="재무제표종류를 선택하세요", style={'margin-bottom': '20px'}),

    html.Label("보고서 종류 선택(*밑에 비교할 보고서를 선택할 경우 이전 보고서만 선택해야함)"),
    dcc.Dropdown(id='independent-report-dropdown', options=report_options, placeholder="보고서를 선택하세요", style={'margin-bottom': '20px'}),

    html.Label("연도 선택"),
    dcc.Dropdown(id='independent-year-dropdown', options=year_options, placeholder="연도를 선택하세요", style={'margin-bottom': '20px'}),

    html.Label("재무제표명 선택 (연결/별도)"),
    dcc.Dropdown(id='independent-statement-type-dropdown', options=statement_type_options, placeholder="연결재무제표 또는 별도재무제표를 선택하세요", style={'margin-bottom': '20px'}),

    html.Label("비교할 보고서 선택 (Optional)"),
    dcc.Dropdown(id='independent-compare-report-dropdown', options=report_options, placeholder="비교할 보고서를 선택하세요", style={'margin-bottom': '20px'}),

    html.Label("등락 비율 선택"),
    dcc.Slider(
        id='independent-rate-change-slider',
        min=min_rate_change,
        max=max_rate_change,
        step=step_rate_change,
        value=min_rate_change,
        marks={i: f'{i}%' for i in range(min_rate_change, max_rate_change + step_rate_change, step_rate_change)},
        tooltip={"placement": "bottom", "always_visible": True},
    ),
    
    html.Div(id='independent-rate-change-results', 
             style={'border': '1px solid black', 'padding': '20px', 'border-radius': '10px', 'margin-top': '20px'}),
    # 캡처 영역을 지정하는 div
    html.Div(id='capture-div', style={'display': 'none'}),
])

def get_stock_price(company_name):
    try:
        # 종목코드 변환
        stock_code = financial_data[financial_data['회사명'] == company_name]['종목코드'].values[0]
        stock_code = str(stock_code).zfill(6)  # 종목코드를 6자리로 맞춤

        # 종목코드에 괄호가 포함되지 않도록 수정
        stock_code = re.sub(r'\[|\]', '', stock_code)

        # 현재 날짜를 구함
        today = datetime.datetime.today().strftime('%Y-%m-%d')

        # fdr 라이브러리를 사용하여 주가 조회
        stock_data = fdr.DataReader(stock_code, today, today)
        if stock_data.empty:
            return "주가 데이터 없음"
        return stock_data['Close'].values[0]
    except IndexError:
        return "종목코드 조회 실패"
    except Exception as e:
        return f"주가 조회 중 오류 발생: {str(e)}"

# 콜백 함수 정의 (선택된 정보와 테이블 모두 처리)
@app.callback(
    [Output('selected-info', 'children'),
     Output('financial-data-table', 'children'),
     Output('download-dataframe-csv', 'data'),
     Output('download-csv-btn', 'n_clicks')],
    [Input('company-dropdown', 'value'),
     Input('report-dropdown', 'value'),
     Input('statement-type-dropdown', 'value'),
     Input('financial-statement-dropdown', 'value'),
     Input('year-dropdown', 'value'),
     Input('download-csv-btn', 'n_clicks')],
    prevent_initial_call=True  # 버튼을 누르기 전에는 콜백이 실행되지 않도록 설정
)

def update_dashboard(selected_company, selected_report, selected_statement_type, selected_financial_statement, selected_year, n_clicks):
    selected_info_output = ""
    financial_table_output = "회사명, 보고서, 재무제표명을 선택해 주세요."
    underline_accounts_income = ["매출액", "영업이익", "당기순이익"]
    underline_accounts_financial = ["유동자산", "비유동자산", "자산총계", "유동부채", "비유동부채", '부채총계', "비지배지분", '지배기업소유주지분', "자본총계"]

    download_data = None  # CSV 다운로드 데이터 기본값 설정

    # 주가는 손익계산서 선택 시에만 가져옴
    stock_price = None
    if selected_company and selected_financial_statement == '손익계산서':
        try:
            stock_price = get_stock_price(selected_company)
            selected_info_output += f"\n{datetime.datetime.today().strftime('%Y-%m-%d')} 기준 주가: {stock_price}원"
        except Exception as e:
            selected_info_output += f"\n주가 조회 중 오류 발생: {str(e)}"

    if selected_company and selected_report and selected_statement_type and selected_financial_statement:
        selected_info_output = f"{selected_company} 주요 재무사항 (단위 : 억(원))"

        if selected_financial_statement == '재무상태표':
            data = financial_data
            keywords = underline_accounts_financial  # 재무상태표 키워드
        elif selected_financial_statement == '손익계산서':
            data = income_statement_data
            keywords = underline_accounts_income  # 손익계산서 키워드
        elif selected_financial_statement == '현금흐름표':
            data = cashflow_data
            keywords = []  # 현금흐름표의 경우 필터링 없이 표시

        filtered_data = data[(data['회사명'] == selected_company) & 
                             (data['보고서종류'] == selected_report) &
                             (data['재무제표명'] == selected_statement_type)]

        # 연도가 선택된 경우 필터링 적용
        if selected_year:
            filtered_data = filtered_data[filtered_data['결산연도'].isin([selected_year, selected_year - 1])]

        current_column = get_columns_by_report_type(selected_report, selected_financial_statement)
        if not current_column:
            return selected_info_output, "해당 보고서 종류에 대한 데이터가 없습니다.", None, n_clicks

        years = [2020, 2021, 2022, 2023, 2024]
        if selected_year:
            years = [selected_year - 1, selected_year]  # 선택된 연도와 그 이전 연도만 표시

        final_table_data = {}

        for index, row in filtered_data.iterrows():
            항목명 = row['항목명'].strip()
            
            # 필터링: 선택된 재무제표 종류에 따라 해당 항목이 키워드 목록에 있는지 확인
            if selected_financial_statement in ['재무상태표', '손익계산서'] and 항목명 not in keywords:
                continue  # 키워드에 해당하지 않으면 스킵

            연도 = row['결산연도']
            value = row[current_column]
            
            if 항목명 in final_table_data:
                if pd.notna(value):
                    final_table_data[항목명][연도] = f"{value:,.0f}"
            else:
                final_table_data[항목명] = {year: '-' for year in years}
                if pd.notna(value):
                    final_table_data[항목명][연도] = f"{value:,.0f}"

        # 테이블 헤더 생성
        table_header = [
            html.Thead(html.Tr([
                html.Th("항목명", style={'border': '1px solid black', 'padding': '6px', 'text-align': 'center', 'font-size': '14px'}),
                *[html.Th(f"{year}연도", style={'border': '1px solid black', 'padding': '6px', 'text-align': 'center', 'font-size': '14px'}) for year in years]
            ]))
        ]

        table_body = []
        csv_data = []

        # 손익계산서인 경우 매출액 -> 영업이익 -> 당기순이익 순서로 정렬
        if selected_financial_statement == '손익계산서':
            ordered_items = ["매출액", "영업이익", "당기순이익"]
            for 항목명 in ordered_items:
                if 항목명 in final_table_data:
                    year_data = final_table_data[항목명]

                    # 스타일 적용 설정
                    is_underline = 항목명 in underline_accounts_income
                    is_bold = is_underline
                    background_color = 'lightyellow' if is_underline else 'none'

                    # 행 생성
                    row_data = [html.Td(
                        항목명,
                        style={
                            'border': '1px solid black',
                            'padding': '6px',
                            'text-decoration': 'underline' if is_underline else 'none',
                            'font-weight': 'bold' if is_bold else 'normal',
                            'background-color': background_color
                        }
                    )]
                    csv_row = [항목명]

                    for i, year in enumerate(years):
                        current_value = year_data.get(year, '-')
                        csv_row.append(current_value)
                        if i > 0:
                            previous_value = year_data.get(years[i-1], '-')
                            if current_value != '-' and previous_value != '-':
                                try:
                                    current_float = float(current_value.replace(',', ''))
                                    previous_float = float(previous_value.replace(',', ''))
                                    change = calculate_change(current_float, previous_float)
                                except ValueError:
                                    change = np.nan
                            else:
                                change = np.nan
                            formatted_change, color = format_change(change)
                            row_data.append(html.Td(
                                f"{current_value} {formatted_change}",
                                style={
                                    'border': '1px solid black',
                                    'padding': '6px',
                                    'text-align': 'right',
                                    'color': color,
                                    'background-color': background_color
                                }
                            ))
                        else:
                            row_data.append(html.Td(
                                current_value,
                                style={
                                    'border': '1px solid black',
                                    'padding': '6px',
                                    'text-align': 'right',
                                    'background-color': background_color
                                }
                            ))

                    table_body.append(html.Tr(row_data))
                    csv_data.append(csv_row)

        # 재무상태표나 기타 재무제표는 기존 정렬 방식 유지
        else:
            for 항목명, year_data in final_table_data.items():
                is_underline = False
                is_bold = False
                background_color = 'none'

                if selected_financial_statement == '재무상태표' and 항목명 in underline_accounts_financial:
                    is_underline = True
                    is_bold = True
                    background_color = 'lightyellow'

                row_data = [html.Td(
                    항목명,
                    style={
                        'border': '1px solid black',
                        'padding': '6px',
                        'text-decoration': 'underline' if is_underline else 'none',
                        'font-weight': 'bold' if is_bold else 'normal',
                        'background-color': background_color
                    }
                )]
                csv_row = [항목명]

                for i, year in enumerate(years):
                    current_value = year_data[year]
                    csv_row.append(current_value)
                    if i > 0:
                        previous_value = year_data[years[i-1]]
                        if pd.notna(current_value) and pd.notna(previous_value):
                            try:
                                current_float = float(current_value.replace(',', ''))
                                previous_float = float(previous_value.replace(',', ''))
                                change = calculate_change(current_float, previous_float)
                            except ValueError:
                                change = np.nan
                        else:
                            change = np.nan
                        formatted_change, color = format_change(change)
                        row_data.append(html.Td(
                            f"{current_value} {formatted_change}",
                            style={
                                'border': '1px solid black',
                                'padding': '6px',
                                'text-align': 'right',
                                'color': color,
                                'background-color': background_color
                            }
                        ))
                    else:
                        row_data.append(html.Td(
                            current_value,
                            style={
                                'border': '1px solid black',
                                'padding': '6px',
                                'text-align': 'right',
                                'background-color': background_color
                            }
                        ))

                table_body.append(html.Tr(row_data))
                csv_data.append(csv_row)

        # 손익계산서 선택 시 주가를 테이블 하단에 추가
        if selected_financial_statement == '손익계산서' and stock_price:
            table_body.append(html.Tr([
                html.Td(
                    f"\n{datetime.datetime.today().strftime('%m월 %d일')} 기준 주가", 
                    style={
                        'border': '1px solid black', 
                        'padding': '6px', 
                        'text-align': 'left', 
                        'font-weight': 'bold',  
                        'background-color': 'lightyellow',  
                        'font-size': '17px'
                    }
                ),
                html.Td(
                    f"{stock_price} 원", 
                    colSpan=len(years), 
                    style={
                        'border': '1px solid black', 
                        'padding': '6px', 
                        'text-align': 'left', 
                        'font-weight': 'bold',  
                        'background-color': 'lightyellow',  
                        'font-size': '17px'
                    }
                )
            ]))

        financial_table_output = html.Table(
            children=table_header + [html.Tbody(table_body)],
            style={'border-collapse': 'collapse', 'width': '100%', 'margin': '20px 0', 'border': '1px solid black'}
        )

        if n_clicks and n_clicks > 0:
            # CSV 파일에 기록할 헤더 (주가 칼럼은 추가하지 않음)
            csv_string = "항목명," + ",".join([f'"{year}"' for year in years]) + "\n"
            
            # 데이터 행 추가
            for row in csv_data:
                formatted_row = [f'"{item}"' if isinstance(item, str) and ',' in item else item for item in row]
                csv_string += ",".join([str(i) for i in formatted_row]) + "\n"

            # 주가 행 추가
            if selected_financial_statement == '손익계산서' and stock_price:
                csv_string += f"{datetime.datetime.today().strftime('%m월 %d일')} 기준 주가, " + "," * (len(years) - 5) + f'{stock_price} 원\n'

            bom = '\ufeff'
            if selected_financial_statement == '손익계산서':
                filename = f"{selected_company} 주요 손익계산서 (단위 : 억(원)).csv"
            else:
                filename = f"{selected_company} 주요 재무사항 (단위 : 억(원)).csv"
            download_data = dict(content=bom + csv_string, filename=filename)
            n_clicks = 0

    return selected_info_output, financial_table_output, download_data, n_clicks

@app.callback(
    Output('independent-rate-change-results', 'children'),
    [Input('independent-year-dropdown', 'value'),
     Input('independent-rate-change-slider', 'value'),
     Input('independent-financial-statement-dropdown', 'value'),
     Input('independent-report-dropdown', 'value'),
     Input('independent-statement-type-dropdown', 'value'),
     Input('independent-compare-report-dropdown', 'value')]  # 비교할 보고서 선택 (Optional)
)
def independent_rate_change_results(selected_year, selected_rate_change, selected_financial_statement, selected_report, selected_statement_type, compare_report):
    print(f"선택된 연도: {selected_year}, 선택된 등락 비율: {selected_rate_change}, 선택된 재무제표: {selected_financial_statement}, 선택된 보고서: {selected_report}, 비교 보고서: {compare_report}, 선택된 재무제표명: {selected_statement_type}")
    
    # 입력값 검증
    if not selected_year or not selected_rate_change or not selected_financial_statement or not selected_report or not selected_statement_type:
        return "연도, 등락 비율, 재무제표종류, 보고서, 재무제표명을 선택하세요."
    
    selected_info = html.Div([
        html.P(f"선택된 연도: {selected_year}"),
        html.P(f"선택된 등락 비율: {selected_rate_change}%"),
        html.P(f"선택된 재무제표: {selected_financial_statement}"),
        html.P(f"선택된 보고서: {selected_report}"),
        html.P(f"비교할 보고서: {compare_report if compare_report else '비교 없음'}"),
        html.P(f"선택된 재무제표명: {selected_statement_type}"),
    ], style={'text-align': 'center'})

    underline_accounts_income = ["매출액", "영업이익", "당기순이익"]
    underline_accounts_financial = ["유동자산", "비유동자산", "자산총계", "유동부채", "비유동부채", '부채총계', "비지배지분", '지배기업소유주지분', "자본총계"]

    if selected_financial_statement == '재무상태표':
        data = financial_data  # 재무상태표 데이터
        keywords = underline_accounts_financial
    elif selected_financial_statement == '손익계산서':
        data = income_statement_data  # 손익계산서 데이터
        keywords = underline_accounts_income
    else:
        data = cashflow_data  # 현금흐름표 데이터
        keywords = []

    # 보고서 순서 설정 (1분기 -> 반기 -> 3분기 -> 사업보고서)
    report_order = ['1분기보고서', '반기보고서', '3분기보고서', '사업보고서']

    # 보고서 순서에 따라 최신 보고서가 더 나중에 나와야 하므로 비교 순서를 조정
    if compare_report and report_order.index(compare_report) > report_order.index(selected_report):
        current_column = get_columns_by_report_type(compare_report, selected_financial_statement)
        previous_column = get_columns_by_report_type(selected_report, selected_financial_statement)
        filtered_data_previous_year = data[(data['보고서종류'] == selected_report) & 
                                           (data['재무제표명'] == selected_statement_type) & 
                                           (data['결산연도'] == selected_year) & 
                                           (data['항목명'].isin(keywords))]
        filtered_data_current_year = data[(data['보고서종류'] == compare_report) & 
                                          (data['재무제표명'] == selected_statement_type) & 
                                          (data['결산연도'] == selected_year) & 
                                          (data['항목명'].isin(keywords))]
    else:
        current_column = get_columns_by_report_type(selected_report, selected_financial_statement)
        previous_column = get_columns_by_report_type(compare_report, selected_financial_statement) if compare_report else current_column
        filtered_data_previous_year = data[(data['보고서종류'] == (compare_report if compare_report else selected_report)) & 
                                           (data['재무제표명'] == selected_statement_type) & 
                                           (data['결산연도'] == (selected_year - 1)) & 
                                           (data['항목명'].isin(keywords))]
        filtered_data_current_year = data[(data['보고서종류'] == selected_report) & 
                                          (data['재무제표명'] == selected_statement_type) & 
                                          (data['결산연도'] == selected_year) & 
                                          (data['항목명'].isin(keywords))]

    if filtered_data_previous_year.empty or filtered_data_current_year.empty:
        return f"{selected_year} 또는 {selected_year-1}에 대한 데이터가 없습니다. 다른 연도를 선택하세요."

    # 등락 비율 구간 설정
    lower_bound = selected_rate_change
    upper_bound = np.inf if selected_rate_change == max_rate_change else selected_rate_change + step_rate_change
    print(f"등락 비율 구간: {lower_bound}% ~ {upper_bound}%")

    result_list = []
    for company in filtered_data_current_year['회사명'].unique():
        current_year_data = filtered_data_current_year[filtered_data_current_year['회사명'] == company]
        previous_year_data = filtered_data_previous_year[filtered_data_previous_year['회사명'] == company]

        if not previous_year_data.empty and not current_year_data.empty:
            for 항목명 in keywords:
                previous_value = previous_year_data[previous_year_data['항목명'] == 항목명][previous_column].values[0] if not previous_year_data[previous_year_data['항목명'] == 항목명].empty else np.nan
                current_value = current_year_data[current_year_data['항목명'] == 항목명][current_column].values[0] if not current_year_data[current_year_data['항목명'] == 항목명].empty else np.nan

                if pd.notna(previous_value) and pd.notna(current_value):
                    try:
                        rate_change = calculate_change(float(current_value), float(previous_value))
                        
                        # 주어진 lower_bound와 upper_bound 범위 내의 항목만 포함
                        if lower_bound <= abs(rate_change) < upper_bound:
                            formatted_change, color = format_change(rate_change)
                            result_list.append({
                                '회사명': company,
                                '항목명': 항목명,
                                f'{selected_report}' if compare_report else f'{selected_year-1} 연도': f"{previous_value:,.0f}",
                                f'{compare_report}' if compare_report else f'{selected_year} 연도': f"{current_value:,.0f}",
                                '변화율': formatted_change,  # 포맷된 변화율 추가
                                'color': color  # 색상 추가
                            })
                    except ValueError as e:
                        print(f"변화율 계산 오류: {e}")
                        continue

    if result_list:
        df = pd.DataFrame(result_list)
        df = df.sort_values(by=['항목명', '변화율'], ascending=[True, False])

        # 테이블 생성
        result_table = dash_table.DataTable(
            data=df.to_dict('records'),
            columns=[
                {"name": "회사명", "id": "회사명"},
                {"name": "항목명", "id": "항목명"},
                {"name": f'{selected_report}' if compare_report else f'{selected_year-1} 연도', "id": f'{selected_report}' if compare_report else f'{selected_year-1} 연도'},
                {"name": f'{compare_report}' if compare_report else f'{selected_year} 연도', "id": f'{compare_report}' if compare_report else f'{selected_year} 연도'},
                {"name": "변화율", "id": "변화율"}
            ],
            sort_action='native',  # 오름차순/내림차순 정렬 가능
            style_cell={'textAlign': 'left', 'padding': '5px'},
            style_data_conditional=[
                # 상승한 항목은 빨간색으로 표시
                {
                    'if': {
                        'filter_query': '{color} = red',
                        'column_id': '변화율'
                    },
                    'color': 'red',
                    'fontWeight': 'bold'
                },
                # 하락한 항목은 파란색으로 표시
                {
                    'if': {
                        'filter_query': '{color} = blue',
                        'column_id': '변화율'
                    },
                    'color': 'blue',
                    'fontWeight': 'bold'
                }
            ],
            style_header={
                'backgroundColor': 'lightgrey',
                'fontWeight': 'bold',
                'border': '1px solid black'
            },
            style_data={
                'border': '1px solid grey'
            },
            page_size=25  # 페이지 당 50개 항목 표시
        )

        # 선택된 정보 텍스트와 결과 테이블을 함께 반환
        return html.Div([selected_info, result_table])
    else:
        return html.Div([selected_info, "선택한 구간에 해당하는 회사가 없습니다."])
    

app.clientside_callback(
    """
    function(n_clicks) {
        if (n_clicks) {
            const element = document.getElementById('table-and-info');
            if (element) {
                html2canvas(element).then(canvas => {
                    const link = document.createElement('a');
                    link.href = canvas.toDataURL();
                    link.download = 'financial_table.png';
                    link.click();
                }).catch(function(error) {
                    console.error('Error capturing the table: ', error);
                });
            } else {
                console.warn('Table element not found.');
            }
        }
        return null;
    }
    """,
    Output('capture-div', 'children'),
    Input('download-btn', 'n_clicks')
)

# 앱 실행
if __name__ == '__main__':
    app.run_server(debug=True, host='0.0.0.0', port=8050)