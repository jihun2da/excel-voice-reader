# 🚀 Streamlit Cloud 배포 완전 가이드

## 📋 **배포 전 체크리스트**

### ✅ **필요한 파일들**
- [x] `streamlit_app.py` - 메인 앱 파일
- [x] `requirements.txt` - 의존성 목록
- [x] `.streamlit/config.toml` - Streamlit 설정
- [x] `.gitignore` - Git 무시 파일
- [x] `README.md` - 프로젝트 설명

### ✅ **GitHub 저장소 준비**
1. GitHub 계정 필요
2. Public 저장소 생성 필요 (무료 배포를 위해)

## 🎯 **단계별 배포 가이드**

### **1단계: GitHub 저장소 생성**

1. **GitHub.com** 접속 후 로그인
2. **"New repository"** 클릭
3. **저장소 정보 입력**:
   - Repository name: `excel-voice-reader`
   - Description: `Excel Voice Reader - Streamlit Web App`
   - Public 선택 (중요!)
   - README, .gitignore, license 체크 해제
4. **"Create repository"** 클릭

### **2단계: 로컬 파일 업로드**

#### **방법 1: GitHub 웹 인터페이스 사용 (추천)**
1. GitHub 저장소 페이지에서 **"uploading an existing file"** 클릭
2. 모든 파일들을 드래그 앤 드롭으로 업로드:
   - `streamlit_app.py`
   - `requirements.txt`
   - `.streamlit/config.toml`
   - `.gitignore`
   - `README.md`
3. **Commit message** 입력: `Initial commit: Excel Voice Reader`
4. **"Commit changes"** 클릭

#### **방법 2: Git 명령어 사용 (Git 설치된 경우)**
```bash
# Git 초기화
git init

# 파일 추가
git add .

# 커밋
git commit -m "Initial commit: Excel Voice Reader Streamlit app"

# 메인 브랜치 설정
git branch -M main

# 원격 저장소 연결 (YOUR_USERNAME을 실제 사용자명으로 변경)
git remote add origin https://github.com/YOUR_USERNAME/excel-voice-reader.git

# GitHub에 푸시
git push -u origin main
```

### **3단계: Streamlit Cloud 배포**

1. **Streamlit Cloud** (https://share.streamlit.io) 접속
2. **GitHub 계정으로 로그인**
3. **"New app"** 버튼 클릭
4. **배포 설정**:
   - **Repository**: `YOUR_USERNAME/excel-voice-reader`
   - **Branch**: `main`
   - **Main file path**: `streamlit_app.py`
   - **App URL**: `excel-voice-reader` (또는 원하는 이름)
5. **"Deploy!"** 클릭

### **4단계: 배포 완료 확인**

- 배포 진행 상황을 실시간으로 확인
- 완료되면 제공되는 URL로 접속 테스트
- 예: `https://excel-voice-reader-YOUR_USERNAME.streamlit.app`

## 🔧 **배포 후 설정**

### **앱 설정 확인**
1. **Streamlit Cloud 대시보드**에서 앱 선택
2. **Settings** 탭에서 설정 확인:
   - **App URL**: 커스텀 URL 설정 가능
   - **Sharing**: 공개/비공개 설정
   - **Resources**: 메모리, CPU 설정

### **도메인 설정 (선택사항)**
- 커스텀 도메인 연결 가능
- SSL 인증서 자동 적용

## 🐛 **문제 해결**

### **배포 실패 시**
1. **requirements.txt** 확인
2. **Python 버전** 호환성 확인
3. **파일 경로** 정확성 확인

### **앱 실행 오류 시**
1. **Streamlit Cloud 로그** 확인
2. **의존성 설치** 상태 확인
3. **코드 문법** 오류 확인

## 📊 **성능 최적화**

### **메모리 사용량**
- Streamlit Cloud 무료 티어: 1GB RAM
- 대용량 파일 처리 시 주의

### **파일 크기 제한**
- 업로드 파일: 최대 200MB
- 임시 파일 자동 정리

## 🔄 **업데이트 배포**

### **코드 수정 후 재배포**
1. GitHub에서 파일 수정
2. Streamlit Cloud에서 **"Reboot app"** 클릭
3. 자동으로 최신 코드로 재배포

### **의존성 업데이트**
1. `requirements.txt` 수정
2. 앱 자동 재시작

## 📈 **모니터링**

### **사용량 확인**
- Streamlit Cloud 대시보드에서 사용량 확인
- 무료 티어 제한: 월 1,000분

### **성능 모니터링**
- 앱 응답 시간
- 메모리 사용량
- 오류 로그

## 🎉 **배포 완료!**

배포가 완료되면:
- ✅ 전 세계 어디서나 접근 가능
- ✅ 모바일/태블릿 지원
- ✅ 자동 HTTPS 적용
- ✅ 무료 호스팅

**배포된 앱 URL**: `https://excel-voice-reader-YOUR_USERNAME.streamlit.app`
