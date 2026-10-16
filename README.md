# 나염·발주 엑셀 자동 가공 웹매크로

## 포함 파일
- `app.py`: 웹 화면 및 엑셀 변환 로직
- `requirements.txt`: 배포 환경에 필요한 Python 패키지

## 로컬 실행 테스트
1. Python 3.10 이상 설치
2. 터미널에서 프로젝트 폴더로 이동
3. `pip install -r requirements.txt`
4. `streamlit run app.py`
5. 브라우저에서 표시되는 주소로 접속

## 인터넷 링크 배포 (Streamlit Community Cloud)
1. GitHub에 새 저장소를 만들고 `app.py`, `requirements.txt` 두 파일을 업로드합니다.
2. https://share.streamlit.io/ 에 접속해 로그인합니다.
3. `Create app`을 눌러 GitHub 저장소, 브랜치, `app.py`를 선택합니다.
4. 배포가 끝나면 `https://....streamlit.app` 형태의 주소가 생성됩니다. 그 링크를 다른 PC에 공유하면 됩니다.

공식 배포 안내: https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy

## 주의
- 기존 매크로와 마찬가지로 입력은 `.xls` 형식입니다.
- 원본 코드의 데이터 범위(11번째 행부터 마지막 행 직전까지) 및 열 위치를 유지했습니다. 실제 발주 파일 샘플로 테스트한 후 운영에 사용하세요.
- 공개 링크로 배포하면 링크를 아는 사람이 접근할 수 있습니다. 거래처 정보 등 민감한 데이터가 포함된다면 접근 제어가 가능한 배포 방식을 고려하세요.
