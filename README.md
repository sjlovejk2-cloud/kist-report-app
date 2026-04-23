# KIST Report App

KIST 시설/계약 실무용 Streamlit 앱 프로젝트입니다.

주요 기능
- AI 보고서 생성 및 HWPX 다운로드
- 설계서 파일 업로드 후 기본 점검
- 설계서 검토 결과를 계산기/계약판단 도구로 연계
- 조달청 제비율 계산기
- 계약 방식 판단기
- KIST 원규·지침 RAG 검색
- 부서정보.md 기반 컨텍스트 반영

프로젝트 주요 파일
- `app.py` : 메인 Streamlit 앱
- `keyword_report.py` : AI 보고서 생성, HWPX 생성, 설계서 AI 검토
- `kist_rag.py` : 원규·지침 벡터DB 검색 모듈
- `build_index.py` : 원규·지침 HWP 파일 인덱싱 스크립트
- `설계검토기준/` : 설계서 검토용 기준 문서
- `kist_db/` : 생성된 ChromaDB 벡터DB
- `부서정보.md` : 보고서 생성 시 반영되는 부서 컨텍스트
- `보고서 초안 양식.hwpx` : HWPX 출력 템플릿

실행 환경
- Windows 권장
- Python 3.12+ 권장
- Streamlit 기반

기본 설치
```bash
pip install -r requirements.txt
```

추가로 필요할 수 있는 패키지
```bash
pip install chromadb pywin32
```

실행 방법
```bash
streamlit run app.py
```

앱에서 사용하는 주요 데이터
- `부서정보.md`
  - AI 보고서 생성 시 자동 반영
- `설계검토기준/`
  - 설계서 검토 기준 문서
- `kist_db/`
  - 원규·지침 검색용 DB

벡터DB 재생성
`build_index.py`는 아래 경로의 HWP 파일을 읽어 DB를 다시 만듭니다.
- 기본 경로: `F:\스마트규정시스템`

실행:
```bash
python build_index.py
```

주의사항
- `build_index.py`는 HWP COM 자동화를 사용하므로 Windows + 한글(HWP) 환경이 적합합니다.
- `kist_db/`를 제거하면 원규·지침 검색 기능이 동작하지 않을 수 있습니다.
- 설계서 자동 추출값은 보조용이므로, 계산기/계약판단 전송 후 반드시 확인이 필요합니다.
- GitHub 업로드 시 개인 문서/대용량 참고자료는 프로젝트 폴더 밖에서 별도 관리하는 것을 권장합니다.

현재 저장소에 포함된 내용
- 앱 실행 코드
- 설계서 검토 기준 문서
- 현재 생성된 `kist_db/`
- 실행/운영 보조 스크립트 및 가이드

빠른 시작 순서
1. 프로젝트 폴더로 이동
2. `pip install -r requirements.txt`
3. 필요 시 `pip install chromadb pywin32`
4. `streamlit run app.py`
5. 원규 검색이 안 되면 `python build_index.py` 재실행

최근 반영 사항
- 설계서 검토 탭 사용 가능 상태로 복구
- 설계서 검토 결과를 계산기/계약판단 도구로 연계하는 1단계 기능 추가
- 설계서 AI 검토 호출 인자 오류 수정
- 자동 금액 추출 로직의 소액 오탐 방지 보정
