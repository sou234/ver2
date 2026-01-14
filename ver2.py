import streamlit as st
import pandas as pd
import plotly.express as px
import FinanceDataReader as fdr
import requests
import urllib3
from io import BytesIO
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import pytz
import feedparser
from etf import ActiveETFMonitor
import yfinance as yf
from curl_cffi import requests as curequests
import re
from collections import Counter
import plotly.graph_objects as go

# 보안 인증서 경고 무시
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 엑셀 다운로드용 함수
def to_excel(df_new, df_inc, df_dec, df_all, date):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_new.to_excel(writer, index=False, sheet_name='신규편입')
        df_inc.to_excel(writer, index=False, sheet_name='비중확대')
        df_dec.to_excel(writer, index=False, sheet_name='비중축소')
        df_all.to_excel(writer, index=False, sheet_name='전체포트폴리오')
    processed_data = output.getvalue()
    return processed_data

# ---------------------------------------------------------
# 1. 페이지 설정
# ---------------------------------------------------------
st.set_page_config(
    page_title="MAS Decision Support System V3.0",
    page_icon="🍊",
    layout="wide"
)

# ---------------------------------------------------------
# 2. 데이터 수집 함수
# ---------------------------------------------------------

@st.cache_data(ttl=600)
def fetch_market_data():
    """시장/거시 지표 수집 (yfinance + curl_cffi)"""
    tickers = {
        "KOSPI": "^KS11", 
        "S&P500": "^GSPC", 
        "USD/KRW": "KRW=X",
        "Bitcoin": "BTC-USD",
        "VIX (공포)": "^VIX",
        "US 10Y (금리)": "^TNX",
        "WTI Rough (유가)": "CL=F",
        "Gold (금)": "GC=F"
    }
    market_data, history_data = {}, {}
    
    # 세션 생성 (봇 탐지 우회)
    session = curequests.Session(impersonate="chrome")
    session.verify = False

    for name, ticker in tickers.items():
        try:
            # yfinance로 데이터 수집
            stock = yf.Ticker(ticker, session=session)
            # 최근 1년치
            df = stock.history(period="1y")
            
            if not df.empty:
                current = df['Close'].iloc[-1]
                # 전일 데이터가 있으면 변동 계산
                if len(df) >= 2:
                    prev = df['Close'].iloc[-2]
                    pct = ((current - prev) / prev * 100)
                    change = current - prev
                else:
                    change = 0
                    pct = 0
                
                # 20일 이평선 트렌드
                df['MA20'] = df['Close'].rolling(window=20).mean()
                ma20 = df['MA20'].iloc[-1] if not pd.isna(df['MA20'].iloc[-1]) else current
                trend = "상승 🐂" if current > ma20 else "하락 🐻"
                
                market_data[name] = {
                    "price": current, 
                    "change": change, 
                    "pct_change": pct, 
                    "trend": trend
                }
                history_data[name] = df
        except Exception as e:
            # print(f"Error fetching {name}: {e}")
            pass
            
    return market_data, history_data

@st.cache_data(ttl=1800)
def fetch_industry_news(topic):
    """구글 뉴스 RSS를 통해 특정 토픽의 뉴스 수집"""
    # 주제별 검색 쿼리 매핑
    queries = {
        "AI & 반도체": "Nvidia OR OpenAI OR TSMC OR Samsung Electronics semiconductor",
        "2차전지 & EV": "Tesla OR CATL OR LG Energy Solution OR electric vehicle battery",
        "바이오 & 헬스케어": "Eli Lilly OR Novo Nordisk OR biotech OR FDA approval",
        "글로벌 거시경제": "Federal Reserve OR inflation OR interest rate OR US economy"
    }
    
    query = queries.get(topic, "Global Economy")
    encoded_query = requests.utils.quote(query)
    # 구글 뉴스 RSS URL (언어: 영어/한국어 섞여있을 수 있음, 여기선 US edition 사용)
    rss_url = f"https://news.google.com/rss/search?q={encoded_query}&hl=en-US&gl=US&ceid=US:en"
    
    try:
        feed = feedparser.parse(rss_url)
        news_items = []
        for entry in feed.entries[:10]: # 최신 10개만
            news_items.append({
                "title": entry.title,
                "link": entry.link,
                "published": entry.published,
                "source": entry.source.title if hasattr(entry, 'source') else "Google News"
            })
        return news_items
    except Exception as e:
        return []

def analyze_news_keywords(news_items):
    """뉴스 제목에서 키워드 추출 및 빈도 분석"""
    text = " ".join([item['title'] for item in news_items])
    # 영문, 숫자만 남기고 제거
    text = re.sub(r'[^a-zA-Z\s]', '', text)
    words = text.lower().split()
    
    # 불용어 처리 (간단한 리스트)
    stop_words = {'to', 'in', 'for', 'of', 'and', 'the', 'a', 'on', 'at', 'with', 'by', 'as', 'is', 'new', 'stocks', 'market'}
    keywords = [w for w in words if w not in stop_words and len(w) > 2]
    
    return Counter(keywords).most_common(10)

# 데이터 로드
metrics, histories = fetch_market_data()

# ---------------------------------------------------------
# 3. 사이드바 구성
# ---------------------------------------------------------
with st.sidebar:
    st.title("🍊 Mirae Asset")
    st.subheader("고객자산배분본부")
    st.caption("Ver 3.0 - Correlation & Comparison")
    st.markdown("---")
    
    menu = st.radio("메뉴 선택", ["📌 시장 동향", "🔍 기업 펀더멘털 스카우터", "📰 글로벌 산업 뉴스", "📊 타임폴리오 ETF 분석"])
    
    if st.button("🔄 데이터 새로고침"):
        st.cache_data.clear()

# ---------------------------------------------------------
# 4. 메인 화면
# ---------------------------------------------------------

if menu == "📌 시장 동향":
    st.title("📈 Global Market Monitor")
    
    # 8개 지표를 4열 2행으로 배치
    row1_cols = st.columns(4)
    row2_cols = st.columns(4)
    
    indicators_row1 = ["KOSPI", "S&P500", "USD/KRW", "Bitcoin"]
    indicators_row2 = ["VIX (공포)", "US 10Y (금리)", "WTI Rough (유가)", "Gold (금)"]
    
    def display_metric(col, key):
        if key in metrics:
            d = metrics[key]
            current_val = d['price']
            
            # 포맷팅 설정 (지수/상품마다 다름)
            if "KRW" in key:
                fmt = "{:,.2f}원"
            elif "VIX" in key or "10Y" in key:
                fmt = "{:,.2f}"
            else:
                fmt = "{:,.2f}"
            
            col.metric(
                label=key,
                value=fmt.format(current_val),
                delta=f"{d['change']:+.2f} ({d['pct_change']:+.2f}%) / {d['trend']}",
                delta_color="inverse" if "KRW" in key or "VIX" in key or "10Y" in key else "normal"
            )
            
            # 미니 차트 (확장기능)
            if key in histories and not histories[key].empty:
                col.line_chart(histories[key]['Close'], height=100)

    # 1열 출력
    for col, key in zip(row1_cols, indicators_row1):
        display_metric(col, key)
        
    st.markdown("---")
        
    # 2열 출력
    for col, key in zip(row2_cols, indicators_row2):
        display_metric(col, key)

    # [새 기능] 자산 상관관계 히트맵
    st.subheader("📊 자산 상관관계 히트맵 (1 Year)")
    st.markdown("주요 자산 간의 **가격 상관계수**를 분석하여 분산 투자 효과를 점검합니다.")
    
    if histories:
        # 데이터 병합
        combined_df = pd.DataFrame()
        
        for key, df in histories.items():
            if not df.empty:
                # 1. Series 추출
                series = df['Close'].copy()
                
                # 2. 인덱스 Timezone 제거 및 날짜로 변환 (서로 다른 타임존/시간대 정렬 문제 해결)
                # DatetimeIndex인지 확인 후 처리
                if isinstance(series.index, pd.DatetimeIndex):
                    series.index = series.index.normalize().tz_localize(None)
                
                # 3. 데이터프레임에 추가 (자동으로 날짜 기준 Join됨)
                combined_df[key] = series
        
        # 결측치 제거 (모든 자산이 거래된 날만 포함 -> 휴장일/주말 제외 효과)
        # 우선 원본 유지하면서 dropna
        corr_df = combined_df.dropna()
        
        if not corr_df.empty and len(corr_df) > 10: # 최소 10일치 이상 데이터 필요
            corr_matrix = corr_df.corr()
            
            fig_corr = px.imshow(corr_matrix, 
                                text_auto='.2f', 
                                aspect="auto",
                                color_continuous_scale="RdBu_r", # +1(상관높음)=Red, -1(역상관)=Blue
                                origin='lower')
            st.plotly_chart(fig_corr, use_container_width=True)
            
            st.caption(f"* 분석 기간: {corr_df.index.min().date()} ~ {corr_df.index.max().date()} ({len(corr_df)} 영업일 기준)")
        else:
            st.warning("상관관계를 계산할 공통 데이터가 충분하지 않습니다. (서로 다른 휴장일/데이터 부족 등)")

elif menu == "🔍 기업 펀더멘털 스카우터":
    st.title("🔍 Stock Fundamental Scout")
    st.markdown("관심 종목의 **핵심 펀더멘털 지표**와 **컨센서스**를 한눈에 파악하세요.")
    
    col1, col2 = st.columns([1, 3])
    with col1:
        ticker_input = st.text_input("티커 입력 (예: NVDA, AAPL, 005930.KS)", "NVDA").strip().upper()
    with col2:
        st.write("") 
        st.write("")
        if st.button("스카우팅 시작"):
            st.session_state['scout_trigger'] = True

        if ticker_input:
            try:
                session = curequests.Session(impersonate="chrome")
                session.verify = False
                stock = yf.Ticker(ticker_input, session=session)
                info = stock.info
                
                # 1. 헤더 정보
                st.subheader(f"{info.get('longName', ticker_input)} ({ticker_input})")
                
                # 가격 정보
                current_price = info.get('currentPrice', info.get('previousClose', 0))
                target_price = info.get('targetMeanPrice', 0)
                
                # 2. 핵심 지표 카드
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("현재 주가", f"${current_price:,.2f}" if current_price else "N/A")
                m2.metric("시가총액", f"${info.get('marketCap', 0)/1e9:,.1f} B" if info.get('marketCap') else "N/A")
                m3.metric("52주 최고가", f"${info.get('fiftyTwoWeekHigh', 0):,.2f}")
                m4.metric("목표주가 (Mean)", f"${target_price:,.2f}" if target_price else "N/A", 
                          delta=f"{(target_price/current_price - 1)*100:.1f}% Upside" if target_price and current_price else None)

                st.markdown("---")
                
                # 3. 상세 펀더멘털 탭
                t1, t2 = st.tabs(["📊 밸류에이션 & 수익성", "📈 주가 차트"])
                
                with t1:
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown("##### 💎 밸류에이션")
                        df_val = pd.DataFrame([
                            {"지표": "Trailing P/E", "값": info.get('trailingPE', 'N/A')},
                            {"지표": "Forward P/E", "값": info.get('forwardPE', 'N/A')},
                            {"지표": "PEG Ratio", "값": info.get('pegRatio', 'N/A')},
                            {"지표": "Price/Book (PBR)", "값": info.get('priceToBook', 'N/A')},
                            {"지표": "Price/Sales (PSR)", "값": info.get('priceToSalesTrailing12Months', 'N/A')},
                        ])
                        st.dataframe(df_val, hide_index=True, use_container_width=True)
                        
                    with c2:
                        st.markdown("##### 💰 수익성 & 배당")
                        df_prf = pd.DataFrame([
                            {"지표": "ROE", "값": f"{info.get('returnOnEquity', 0)*100:.2f}%" if info.get('returnOnEquity') else 'N/A'},
                            {"지표": "Profit Margin", "값": f"{info.get('profitMargins', 0)*100:.2f}%" if info.get('profitMargins') else 'N/A'},
                            {"지표": "Dividend Yield", "값": f"{info.get('dividendRate', 0)*100:.2f}%" if info.get('dividendRate') else 'N/A'},
                            {"지표": "Beta", "값": info.get('beta', 'N/A')},
                        ])
                        st.dataframe(df_prf, hide_index=True, use_container_width=True)
                    
                    st.info(f"💡 {info.get('longBusinessSummary', '기업 설명 정보가 없습니다.')[:300]}...")

                with t2:
                    st.markdown("##### 최근 1년 주가 흐름")
                    hist = stock.history(period="1y")
                    if not hist.empty:
                        st.line_chart(hist['Close'])
                    else:
                        st.warning("주가 데이터를 불러올 수 없습니다.")
                        
            except Exception as e:
                st.error(f"데이터를 가져오는 중 오류가 발생했습니다: {e}") 


elif menu == "📰 글로벌 산업 뉴스":
    st.title("📰 Global Industry & Macro News")
    st.markdown("주요 산업 및 거시 경제 관련 최신 뉴스를 실시간으로 확인하세요.")
    
    # 탭으로 분야 구분
    topics = ["AI & 반도체", "2차전지 & EV", "바이오 & 헬스케어", "글로벌 거시경제"]
    tabs = st.tabs(topics)
    
    for i, topic in enumerate(topics):
        with tabs[i]:
            st.subheader(f"{topic} 주요 뉴스")
            news_items = fetch_industry_news(topic)
            
            if news_items:
                # [새 기능] 키워드 트렌드 분석
                st.markdown("##### ☁️ 뉴스 키워드 트렌드")
                keywords_data = analyze_news_keywords(news_items)
                
                if keywords_data:
                    kw_df = pd.DataFrame(keywords_data, columns=['키워드', '빈도'])
                    fig_kw = px.bar(kw_df, x='키워드', y='빈도', color='빈도', 
                                   title=f"'{topic}' 관련 뉴스 최다 빈출 단어",
                                   color_continuous_scale='Teal')
                    st.plotly_chart(fig_kw, use_container_width=True)
                
                st.markdown("---")
                # 뉴스 리스트
                for item in news_items:
                    with st.container():
                        st.markdown(f"### [{item['title']}]({item['link']})")
                        st.caption(f"{item['source']} | {item['published']}")
                        st.markdown("---")
            else:
                st.info("뉴스를 불러올 수 없습니다.")

elif menu == "📊 타임폴리오 ETF 분석": # 메뉴명 변경
    st.title("📊 TIMEFOLIO ETF Comparison & Monitor")
    
    etf_categories = {
        "해외주식형 (10종)": {
            "글로벌탑픽": "22", "글로벌바이오": "9", "우주테크&방산": "20",
            "S&P500": "5", "나스닥100": "2", "글로벌AI": "6",
            "차이나AI": "19", "미국배당다우존스": "18",
            "미국나스닥100채권혼합50": "10", "글로벌소비트렌드": "8"
        },
        "국내주식형 (7종)": {
            "K신재생에너지": "16", "K바이오": "13", "Korea플러스배당": "12",
            "코스피": "11", "코리아밸류업": "15", "K이노베이션": "17", "K컬처": "1"
        }
    }

    # 분석 모드 선택
    mode = st.radio("분석 모드", ["단일 상품 모니터링", "⚔️ ETF 비교 분석"], horizontal=True)

    if mode == "단일 상품 모니터링":
        c1, c2 = st.columns(2)
        with c1:
            cat = st.selectbox("분류", list(etf_categories.keys()))
        with c2:
            name = st.selectbox("상품명", list(etf_categories[cat].keys()))
        
        target_idx = etf_categories[cat][name]
        
        if st.button("데이터 분석 및 리밸런싱 요약"):
            with st.spinner(f"'{name}' 데이터를 수집 및 분석 중입니다..."):
                try:
                    monitor = ActiveETFMonitor(url=f"https://timefolioetf.co.kr/m11_view.php?idx={target_idx}", etf_name=name)
                    today = datetime.now(pytz.timezone('Asia/Seoul')).strftime("%Y-%m-%d")
                    df_today = monitor.get_portfolio_data(today)
                    monitor.save_data(df_today, today)
                    
                    try:
                        prev_day = monitor.get_previous_business_day(today)
                        df_prev = monitor.load_data(prev_day)
                        analysis = monitor.analyze_rebalancing(df_today, df_prev, prev_day, today)
                        analysis_success = True
                    except Exception as e:
                        st.warning(f"전일 데이터를 찾을 수 없어 리밸런싱 분석을 건너뜁니다: {e}")
                        analysis_success = False
                        df_prev = None

                    st.success(f"✅ {name} 데이터 분석 완료" + (f" (기준: {today} vs {prev_day})" if analysis_success else ""))

                    if analysis_success:
                        st.subheader("🔄 리밸런싱 정밀 분석 (시장수익률 조정 반영)")
                        m1, m2, m3, m4 = st.columns(4)
                        m1.metric("비중 확대", f"{len(analysis['increased_stocks'])} 종목")
                        m2.metric("비중 축소", f"{len(analysis['decreased_stocks'])} 종목")
                        m3.metric("신규 편입", f"{len(analysis['new_stocks'])} 종목")
                        m4.metric("완전 편출", f"{len(analysis['removed_stocks'])} 종목")

                        tab1, tab2, tab3 = st.tabs(["주요 변경내역", "세부 변동", "전체 포트폴리오"])
                        with tab1:
                            c1, c2 = st.columns(2)
                            with c1:
                                st.markdown("##### 🟢 신규 편입")
                                if analysis['new_stocks']:
                                    rows = []
                                    for s in analysis['new_stocks']:
                                        rows.append({"종목명": s['종목명'], "현재비중": f"{s['비중_today']:.2f}%", "순수변동": f"+{s['순수_비중변화']:.2f}%p"})
                                    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
                                else:
                                    st.caption("신규 편입 종목 없음")
                            with c2:
                                st.markdown("##### 🔴 완전 편출")
                                if analysis['removed_stocks']:
                                    rows = []
                                    for s in analysis['removed_stocks']:
                                        rows.append({"종목명": s['종목명'], "이전비중": f"{s['비중_prev']:.2f}%", "순수변동": f"{s['순수_비중변화']:.2f}%p"})
                                    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
                                else:
                                    st.caption("완전 편출 종목 없음")

                        with tab2:
                            c1, c2 = st.columns(2)
                            with c1:
                                st.markdown("##### 🔼 비중 확대 (Top 5)")
                                if analysis['increased_stocks']:
                                    df_inc = pd.DataFrame(analysis['increased_stocks']).sort_values('순수_비중변화', ascending=False).head(5)
                                    display_df = df_inc[['종목명', '비중_prev', '비중_today', '순수_비중변화']].copy()
                                    display_df.columns = ['종목명', '이전(%)', '현재(%)', '변동(%p)']
                                    st.dataframe(display_df.style.format({'이전(%)': '{:.2f}', '현재(%)': '{:.2f}', '변동(%p)': '+{:.2f}'}), hide_index=True, use_container_width=True)
                                else:
                                    st.caption("비중 확대 종목 없음")

                            with c2:
                                st.markdown("##### 🔽 비중 축소 (Top 5)")
                                if analysis['decreased_stocks']:
                                    df_dec = pd.DataFrame(analysis['decreased_stocks']).sort_values('순수_비중변화', ascending=True).head(5)
                                    display_df = df_dec[['종목명', '비중_prev', '비중_today', '순수_비중변화']].copy()
                                    display_df.columns = ['종목명', '이전(%)', '현재(%)', '변동(%p)']
                                    st.dataframe(display_df.style.format({'이전(%)': '{:.2f}', '현재(%)': '{:.2f}', '변동(%p)': '{:.2f}'}), hide_index=True, use_container_width=True)
                                else:
                                    st.caption("비중 축소 종목 없음")
                    else:
                        st.subheader("📋 전체 포트폴리오 구성")

                    col_chart, col_list = st.columns([1, 1])
                    with col_chart:
                        chart_df = df_today.copy()
                        chart_df['비중'] = pd.to_numeric(chart_df['비중'], errors='coerce')
                        chart_df.loc[chart_df['비중'] < 1.0, '종목명'] = '기타'
                        fig = px.pie(chart_df, values="비중", names="종목명", hole=0.4, title="포트폴리오 비중", color_discrete_sequence=px.colors.qualitative.Set3)
                        st.plotly_chart(fig, use_container_width=True)

                    with col_list:
                        top_df = df_today[['종목명', '비중', '수량']].head(15)
                        st.dataframe(top_df.style.format({'비중': '{:.2f}%', '수량': '{:,}'}), use_container_width=True)

                    # 엑셀 다운로드
                    st.markdown("---")
                    e_new = pd.DataFrame(analysis['new_stocks']) if analysis_success and analysis['new_stocks'] else pd.DataFrame()
                    e_inc = pd.DataFrame(analysis['increased_stocks']) if analysis_success and analysis['increased_stocks'] else pd.DataFrame()
                    e_dec = pd.DataFrame(analysis['decreased_stocks']) if analysis_success and analysis['decreased_stocks'] else pd.DataFrame()
                    excel_data = to_excel(e_new, e_inc, e_dec, df_today, today)
                    st.download_button(label="📊 엑셀 리포트 내려받기 (.xlsx)", data=excel_data, file_name=f"{name}_Report_{today}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                except Exception as e:
                    st.error(f"오류 발생: {e}")
    
    # [새 기능] ETF 비교 분석
    elif mode == "⚔️ ETF 비교 분석":
        st.subheader("⚔️ ETF Portfolio Comparison")
        st.markdown("두 개의 Timefolio Active ETF 구성을 비교하여 **교집합 종목**과 **Overlap 비중**을 확인합니다.")
        
        flat_etfs = {}
        for cat, items in etf_categories.items():
            for name, idx in items.items():
                flat_etfs[f"[{cat}] {name}"] = idx
        
        c1, c2 = st.columns(2)
        with c1:
            etf_a_key = st.selectbox("ETF A 선택", list(flat_etfs.keys()), index=0)
        with c2:
            etf_b_key = st.selectbox("ETF B 선택", list(flat_etfs.keys()), index=1)
            
        if st.button("비교 분석 실행"):
            if etf_a_key == etf_b_key:
                st.warning("서로 다른 ETF를 선택해주세요.")
            else:
                with st.spinner("두 ETF 데이터를 수집 및 비교 중..."):
                    try:
                        today = datetime.now(pytz.timezone('Asia/Seoul')).strftime("%Y-%m-%d")
                        
                        # Data A
                        mon_a = ActiveETFMonitor(url=f"https://timefolioetf.co.kr/m11_view.php?idx={flat_etfs[etf_a_key]}")
                        df_a = mon_a.get_portfolio_data(today)
                        
                        # Data B
                        mon_b = ActiveETFMonitor(url=f"https://timefolioetf.co.kr/m11_view.php?idx={flat_etfs[etf_b_key]}")
                        df_b = mon_b.get_portfolio_data(today)
                        
                        # 비교 로직 using 종목코드
                        # 현금 제외
                        df_a = df_a[df_a['종목명'] != '현금']
                        df_b = df_b[df_b['종목명'] != '현금']
                        
                        merged = pd.merge(df_a[['종목코드', '종목명', '비중']], df_b[['종목코드', '종목명', '비중']], 
                                        on='종목코드', how='inner', suffixes=('_A', '_B'))
                        merged['종목명'] = merged['종목명_A'] # 이름 통일
                        
                        # Overlap Weight 계산 (두 비중 중 작은 값의 합)
                        merged['Overlap'] = merged[['비중_A', '비중_B']].min(axis=1)
                        total_overlap = merged['Overlap'].sum()
                        
                        # 결과 표시
                        st.markdown("---")
                        res_col1, res_col2 = st.columns(2)
                        with res_col1:
                            st.metric("공통 보유 종목 수", f"{len(merged)} 개")
                        with res_col2:
                            st.metric("Overlap Weight (중복 비중)", f"{total_overlap:.2f}%")
                            
                        # 시각화 (양쪽 비중 비교)
                        if not merged.empty:
                            st.subheader("📊 공통 종목 비중 비교")
                            merged_sorted = merged.sort_values('Overlap', ascending=False)
                            
                            fig = go.Figure(data=[
                                go.Bar(name=etf_a_key, x=merged_sorted['종목명'], y=merged_sorted['비중_A']),
                                go.Bar(name=etf_b_key, x=merged_sorted['종목명'], y=merged_sorted['비중_B'])
                            ])
                            fig.update_layout(barmode='group', title="공통 종목 비중 비교")
                            st.plotly_chart(fig, use_container_width=True)
                            
                            st.dataframe(merged[['종목명', '비중_A', '비중_B', 'Overlap']].style.format("{:.2f}%", subset=['비중_A', '비중_B', 'Overlap']), use_container_width=True)
                        else:
                            st.info("두 ETF 간 겹치는 종목이 없습니다.")
                            
                    except Exception as e:
                        st.error(f"비교 분석 중 오류 발생: {e}")

    st.markdown("---")
    st.caption("Data source: TIMEFOLIO ETF Official Website")