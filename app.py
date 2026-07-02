import streamlit as st
import pandas as pd
import requests
import time
from io import BytesIO
import openpyxl # 엑셀 서식 유지하면서 셀만 수정하기 위해 추가됨

# --- 페이지 기본 설정 ---
st.set_page_config(page_title="단가 검토 자동화 봇", page_icon="🛒", layout="wide")

# --- 네이버 쇼핑 API 호출 함수 ---
def search_naver_shopping(query, client_id, client_secret):
    url = "https://openapi.naver.com/v1/search/shop.json"
    headers = {
        "X-Naver-Client-Id": client_id,
        "X-Naver-Client-Secret": client_secret
    }
    # 검색어(query)로 1개(display=1)의 결과만 정확도순(sim)으로 가져옴
    params = {"query": query, "display": 1, "sort": "sim"}
    
    try:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status() # 에러 나면 예외 발생
        data = response.json()
        
        if data['items']:
            item = data['items'][0]
            # HTML 태그 제거 (<b> 등)
            title = item['title'].replace('<b>', '').replace('</b>', '')
            return {
                "검색된 상품명": title,
                "최저가(원)": int(item['lprice']),
                "쇼핑몰": item['mallName'],
                "링크": item['link']
            }
        else:
            return {"검색된 상품명": "검색 결과 없음", "최저가(원)": 0, "쇼핑몰": "-", "링크": "-"}
            
    except Exception as e:
        return {"검색된 상품명": f"에러: {e}", "최저가(원)": 0, "쇼핑몰": "-", "링크": "-"}

# --- 메인 화면 UI ---
st.title("🛒 네이버 쇼핑 단가 검토 자동화")
st.markdown("""
**사용 방법:**
1. 검토할 **엑셀 양식 파일(.xlsx)**을 올려라.
2. 실행 버튼을 누르면 원본 양식 그대로 **G열(상품가)**과 **I열(링크)**에 데이터가 채워진데이!
""")

# --- 사이드바: API 키 설정 (보안상 st.secrets 사용) ---
with st.sidebar:
    st.header("🔑 API 키 설정")
    
    # st.secrets에 키가 등록되어 있으면 자동으로 불러오고 입력창 숨김!
    try:
        client_id = st.secrets["NAVER_CLIENT_ID"]
        client_secret = st.secrets["NAVER_CLIENT_SECRET"]
        st.success("✅ 서버 금고에서 API 키를 자동으로 삭 불러왔데이! (입력할 필요 없음)")
    except:
        st.warning("⚠️ 서버 금고에 등록된 키가 없데이. 아래에 직접 입력해라.")
        client_id = st.text_input("Client ID", type="password")
        client_secret = st.text_input("Client Secret", type="password")

# --- 엑셀 파일 업로드 ---
uploaded_file = st.file_uploader("엑셀 파일(.xlsx)을 여기다 던져라", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 화면에 엑셀 데이터 살짝 보여주기 (니 양식 헤더가 3줄이라서 4번째 줄부터 읽게 하려면 header=2)
        df_preview = pd.read_excel(uploaded_file, header=2)
             
        st.subheader("📊 니가 올린 엑셀 데이터 (미리보기)")
        st.dataframe(df_preview.head())
        
        if st.button("🚀 최저가 긁어오기 시작!"):
            if not client_id or not client_secret:
                st.error("마! API 키(Client ID, Secret)가 없으면 조회를 몬한데이!")
            else:
                progress_text = "열심히 검색 중이데이... 쫌만 기다리라..."
                
                # 업로드된 파일의 스트림 위치를 처음으로 되돌림 (pandas에서 한 번 읽었으니까)
                uploaded_file.seek(0)
                
                # openpyxl로 원본 엑셀의 껍데기(서식)를 그대로 유지한 채로 파일 열기
                wb = openpyxl.load_workbook(uploaded_file)
                ws = wb.active
                
                # 니 양식 보니까 실제 데이터가 4번째 줄부터 시작하더라
                total_items = ws.max_row - 3
                my_bar = st.progress(0, text=progress_text)
                
                count = 0
                for row in range(4, ws.max_row + 1):
                    count += 1
                    
                    # C열(상품명), D열(규격), E열(제조사) 데이터 가져오기
                    name = ws[f'C{row}'].value
                    spec = ws[f'D{row}'].value
                    maker = ws[f'E{row}'].value
                    
                    # 상품명이 비어있으면 볼 것도 없이 패스
                    if not name:
                        continue
                        
                    # 검색어 조합 (C+D+E 합치기, 빈칸이면 무시)
                    query_parts = [str(x).strip() for x in [name, spec, maker] if x and str(x).strip() != 'nan' and str(x).strip() != 'None']
                    query = " ".join(query_parts)
                    
                    if not query:
                        continue
                        
                    # 뚝딱 만든 검색어로 네이버 찔러보기
                    search_result = search_naver_shopping(query, client_id, client_secret)
                    
                    # G열(상품가)에 최저가 꽂아넣기
                    if search_result["최저가(원)"] > 0:
                        ws[f'G{row}'].value = search_result["최저가(원)"]
                    else:
                        ws[f'G{row}'].value = "결과 없음"
                        
                    # I열(링크)에 최저가 링크 꽂아넣기
                    if search_result["링크"] != "-":
                        ws[f'I{row}'].value = search_result["링크"]
                        
                    # API 뻗지 말라고 0.1초씩 숨 고르기
                    time.sleep(0.1)
                    
                    # 진행률 쫙쫙 올려주기
                    if total_items > 0:
                        my_bar.progress(count / total_items, text=f"[{query}] 검색 완료! ({count}/{total_items})")
                
                st.success("마! 니 엑셀에 최저가 다 채워넣었다! 밑에 버튼 눌러서 다운 받아라!")
                
                # 엑셀 파일 바이너리로 변환해서 다운로드 준비
                output = BytesIO()
                wb.save(output)
                processed_data = output.getvalue()
                
                st.download_button(
                    label="📥 원본 서식 그대로! 최저가 채워진 엑셀 다운로드",
                    data=processed_data,
                    file_name="단가검토완료.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"엑셀 처리하다가 에러 났다. 엑셀 파일이 열려있거나 양식이 좀 다른갑다: {e}")
