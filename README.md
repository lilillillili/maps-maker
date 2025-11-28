# 회원사 주소 검색 후 지도 만들기

## 🚀 설명
- member_visit.py 파일에서 회원사 목록.xlsx의 회원사명 열에서 주소와 홈페이지 url을 검색, 엑셀로 출력
- member_visit_view (upload).py 에서 489행에 구글 api키 입력 후 html로 지도 출력

## 🛠️ 기술 스택
- Python 3.7

## 🎯 설치 및 실행 방법
1. Python 라이브러리 설치:
   `pip install pandas openpyxl requests beautifulsoup4 lxml`
2. 회원사 목록.xlsx에 회원사명 기입 후 파이썬 스크립트 실행:
   `member_visit.py`
3. 출력된 엑셀 확인:
   `회원사 목록_업데이트.xlsx`
4. 구글 API 키 발급
   `console.cloud.google.com`
5. 489행에 발급받은 키 입력
6. 파이썬 실행:
   `member_visit_view_upload.py`
