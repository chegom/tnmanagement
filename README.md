# KITT 엑셀 변환 시스템

회사 내부용 엑셀 파일 자동 변환 도구

## 🚀 빠른 시작

### 1. 환경 설정
```bash
# Python 가상환경 생성 (권장)
python -m venv venv

# 가상환경 활성화
# Windows:
venv\Scripts\activate
# Mac/Linux:
source venv/bin/activate

# 필수 패키지 설치
pip install -r requirements.txt
```

### 2. 실행
```bash
streamlit run app.py
```

브라우저에서 자동으로 http://localhost:8501 열림

## 📁 프로젝트 구조
```
shopp/
├── app.py                 # Streamlit 메인 앱
├── requirements.txt       # Python 패키지 목록
├── excel_processor.py     # 엑셀 변환 로직
└── README.md             # 이 파일
```

## 🛠 기술 스택
- **Streamlit**: 웹 UI
- **Pandas**: 엑셀 데이터 처리
- **openpyxl**: Excel 파일 읽기/쓰기

## 📝 사용 방법
1. 웹 페이지에서 엑셀 파일 업로드
2. 자동으로 변환 규칙 적용
3. 변환된 파일 다운로드
