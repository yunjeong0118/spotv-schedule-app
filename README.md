# SPOTV 편성표 → 업로드파일 변환 웹 앱

편성표(.xlsx)를 업로드파일 형식으로 변환하고, 기존/변경 편성표를 비교해 변경구간을 뽑고 변경된 셀만 노란색으로 표시한 업로드파일을 만들어주는 도구입니다.

이 도구는 AI를 호출하지 않는 규칙 기반 로직입니다(테두리·셀 색상·도형 텍스트 분석). 따라서 API 키가 필요 없고, 실행 비용도 들지 않습니다.

## 파일 구성
- `app.py` — Streamlit 웹 앱 (화면)
- `schedule_parser.py` — 실제 파싱/비교/엑셀 생성 로직 (핵심)
- `requirements.txt` — 필요한 파이썬 패키지

## 로컬에서 바로 실행해보기

```bash
pip install -r requirements.txt
streamlit run app.py
```

브라우저에서 `http://localhost:8501` 이 자동으로 열립니다.

## 팀 전체가 쓸 수 있게 배포하는 방법

### 방법 A — Streamlit Community Cloud (가장 쉬움, 무료)
1. 이 폴더를 GitHub 저장소(private 가능)에 올립니다.
2. https://share.streamlit.io 에 접속해 GitHub 계정으로 로그인합니다.
3. "New app" → 방금 올린 저장소 선택 → `app.py` 지정 → Deploy.
4. 몇 분 뒤 `https://xxxx.streamlit.app` 형태의 링크가 생깁니다. 이 링크를 팀원에게 공유하면 끝입니다.
5. 사내 문서에만 접근을 제한하고 싶다면 Streamlit Cloud의 "Viewer 인증" 옵션이나, 사내 SSO를 지원하는 유료 플랜(Streamlit in Snowflake 등)을 검토하세요. 방송 편성 데이터가 민감하다면 이 방식보다는 아래 사내 서버 방식을 권장합니다.

### 방법 B — 사내 서버/사내 클라우드에 직접 배포 (보안 데이터에 권장)
사내에 리눅스 서버(또는 사내 클라우드 VM)가 있다면:
```bash
pip install -r requirements.txt
streamlit run app.py --server.port 8080 --server.address 0.0.0.0
```
그 뒤 사내망에서만 접근되는 주소(`http://내부서버주소:8080`)를 팀원에게 공유합니다. 상시 구동을 원하면 `systemd` 서비스나 `pm2`, `nohup` 등으로 백그라운드 실행하거나, 아래 Docker 방식을 사용하세요.

### 방법 C — Docker로 배포
```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY . .
RUN pip install -r requirements.txt
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```
```bash
docker build -t spotv-schedule-app .
docker run -p 8501:8501 spotv-schedule-app
```
사내 컨테이너 레지스트리/오케스트레이션(예: AWS ECS, 사내 쿠버네티스)에 그대로 올릴 수 있습니다.

## 로직을 수정/보강하고 싶을 때
모든 판정 규칙(시간축 계산, 테두리 기반 항목 경계 판정, 색상 기반 방송구분, 자막/해설/수어 도형 매칭)은 `schedule_parser.py` 안의 함수 단위로 분리되어 있습니다. 편성표 서식이 바뀌거나 새로운 규칙이 필요하면 이 파일만 수정하면 되고, `app.py`(화면)는 건드릴 필요가 없습니다.

## 데이터 보안 참고
업로드된 파일은 이 웹 앱 프로세스 메모리에서만 처리되고 디스크에 저장하지 않습니다(Streamlit의 `file_uploader`는 메모리 버퍼로 제공됩니다). 다만 어디에 호스팅하든 사내 정책에 맞는 접근 제어(사내망 한정, 로그인 등)를 반드시 함께 적용하세요.
