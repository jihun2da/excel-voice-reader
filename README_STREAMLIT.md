# 🎵 엑셀 음성 리더 - Streamlit 버전

## 🚀 Streamlit Cloud 배포 가이드

### **1단계: GitHub 저장소 생성**

1. **GitHub에 로그인**하고 새 저장소 생성
2. **저장소 이름**: `excel-voice-reader` (또는 원하는 이름)
3. **Public** 저장소로 설정 (Streamlit Cloud 무료 배포를 위해)

### **2단계: 로컬 파일들을 GitHub에 업로드**

```bash
# Git 초기화 (Git이 설치되어 있다면)
git init
git add .
git commit -m "Initial commit: Excel Voice Reader Streamlit app"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/excel-voice-reader.git
git push -u origin main
```

### **3단계: Streamlit Cloud 배포**

1. **Streamlit Cloud** (https://share.streamlit.io) 접속
2. **GitHub 계정으로 로그인**
3. **"New app"** 클릭
4. **저장소 선택**: `YOUR_USERNAME/excel-voice-reader`
5. **Branch**: `main`
6. **Main file path**: `streamlit_app.py`
7. **"Deploy!"** 클릭

### **4단계: 배포 완료**

- 배포가 완료되면 Streamlit Cloud에서 제공하는 URL로 접속 가능
- 예: `https://excel-voice-reader-YOUR_USERNAME.streamlit.app`

## 📁 **프로젝트 구조**

```
excel-voice-reader/
├── streamlit_app.py          # 메인 Streamlit 앱
├── requirements.txt          # Python 의존성
├── .streamlit/
│   └── config.toml          # Streamlit 설정
├── .gitignore               # Git 무시 파일
├── README.md               # 프로젝트 설명
└── README_STREAMLIT.md     # 배포 가이드
```

## 🛠️ **로컬 실행**

```bash
# 의존성 설치
pip install -r requirements.txt

# Streamlit 앱 실행
streamlit run streamlit_app.py
```

## 🌟 **주요 기능**

- **📂 엑셀 파일 업로드**: .xlsx, .xls 파일 지원
- **🎵 고품질 TTS**: Edge TTS와 브라우저 TTS 지원
- **🎛️ 고급 설정**: 음성, 속도, 엔진 선택
- **📊 실시간 진행률**: 현재 진행 상황 표시
- **⌨️ 키보드 단축키**: 편리한 조작
- **🔍 네이버 검색**: 상품 정보 검색
- **⚡ 자동 진행**: 자동으로 다음 행으로 이동

## 🎯 **사용법**

1. **파일 업로드**: 엑셀 파일을 드래그 앤 드롭
2. **설정 조정**: 사이드바에서 TTS 설정
3. **읽기 시작**: '시작' 버튼 클릭
4. **자동 진행**: 필요시 자동 진행 활성화

## 🔧 **기술 스택**

- **Frontend**: Streamlit
- **TTS**: Edge-TTS, Web Speech API
- **파일 처리**: openpyxl
- **배포**: Streamlit Cloud

## 📝 **라이선스**

MIT License
