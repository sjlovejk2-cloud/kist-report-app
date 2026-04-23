# 초보자용 서버 컴퓨터 Dooray 점검 환경 구축 가이드

## 문서 목적
이 문서는 "코딩 환경이 아무것도 없는 Windows 서버 컴퓨터"에
Dooray 프로젝트 점검용 스크립트를 설치하고 실행할 수 있게 만드는 초보자용 설치 가이드다.

이 가이드의 목표는 아래 2가지다.
1. 서버에서 최신 Dooray 업무를 조회할 수 있게 만들기
2. 보완이 필요한 업무를 리포트로 뽑을 수 있게 만들기

현재 단계에서 하는 일은 "자동 처리 시스템 구축"이 아니라
"프로젝트/업무 보완 점검 환경 구축"이다.

---

## 최종적으로 서버에서 하게 될 일
서버 컴퓨터는 아래 작업만 수행하면 된다.
1. Dooray API Token으로 Dooray 업무 조회
2. 점검 스크립트 실행
3. 결과 파일 저장
4. 나중에는 Dooray 메신저로 자동 발송까지 확장 가능

---

## 준비물
서버 컴퓨터에서 필요한 것
1. Windows 운영체제
2. 인터넷 연결
3. PowerShell 실행 가능
4. Dooray API Token
5. 작업 폴더를 둘 드라이브 공간
   - 권장: F: 드라이브

중요
- 브라우저에서 Dooray 로그인 상태일 필요 없음
- 두레이 웹사이트를 켜둘 필요 없음
- API Token만 있으면 됨

---

## 설치 순서 한눈에 보기
아주 쉽게 보면 아래 순서다.
1. 작업 폴더 만들기
2. Python 설치
3. 필요한 파일 복사
4. PowerShell 실행 정책/환경 변수 설정
5. 테스트 실행
6. 결과 파일 확인
7. 나중에 작업 스케줄러 등록

---

## 1단계. 작업 폴더 만들기
권장 경로
- F:\telegram-deadline-bot\telegram-deadline-bot

만드는 방법
1. 파일 탐색기 열기
2. F: 드라이브로 이동
3. telegram-deadline-bot 폴더 생성
4. 그 안에 다시 telegram-deadline-bot 폴더 생성

최종 경로 예시
- F:\telegram-deadline-bot\telegram-deadline-bot

이 경로를 쓰는 이유
- 기존 작업 파일 경로와 맞추기 쉬움
- 스크립트 수정 없이 그대로 쓰기 좋음

---

## 2단계. Python 설치
이 서버에는 개발 환경이 없다고 했으므로 Python부터 설치해야 한다.

권장 버전
- Python 3.14

설치 방법
1. 웹 브라우저 열기
2. Python 공식 사이트 접속
   - https://www.python.org/downloads/windows/
3. Windows용 Python 3.14 설치 파일 다운로드
4. 설치 프로그램 실행
5. 설치 화면에서 반드시 아래 체크
   - Add python.exe to PATH
6. 설치 진행
7. 가능하면 기본 경로 대신 아래처럼 간단한 경로 권장
   - C:\Python314

설치 후 확인 방법
1. PowerShell 열기
2. 아래 명령 실행

```powershell
python --version
```

또는

```powershell
C:\Python314\python.exe --version
```

정상 예시
- Python 3.14.x

주의
- 현재 스크립트는 C:\Python314\python.exe 경로를 우선 사용하도록 되어 있음
- 따라서 가능하면 실제 설치 경로도 C:\Python314로 맞추는 것이 가장 편함

---

## 3단계. 필요한 파일 복사
서버에 아래 파일들을 복사해야 한다.

필수 파일
- validate_preconstruction_meetings.py
- preconstruction_validation_rules.yaml
- run_latest_preconstruction_validation.ps1

있으면 좋은 참고 파일
- daily_deadline_report.ps1
- server_dooray_check_setup_guide.md
- server_beginner_dooray_setup_guide.md

복사 위치
- F:\telegram-deadline-bot\telegram-deadline-bot

확인 방법
파일 탐색기에서 아래 파일이 보이면 됨
- F:\telegram-deadline-bot\telegram-deadline-bot\validate_preconstruction_meetings.py
- F:\telegram-deadline-bot\telegram-deadline-bot\preconstruction_validation_rules.yaml
- F:\telegram-deadline-bot\telegram-deadline-bot\run_latest_preconstruction_validation.ps1

---

## 4단계. PyYAML 설치
파이썬 스크립트는 PyYAML 패키지가 필요하다.
초기 서버에는 없을 가능성이 높다.

설치 방법
1. PowerShell 열기
2. 아래 명령 실행

```powershell
C:\Python314\python.exe -m pip install pyyaml
```

만약 python 명령이 잘 잡히면 아래도 가능

```powershell
python -m pip install pyyaml
```

설치 확인 방법
아래 명령 실행

```powershell
C:\Python314\python.exe -c "import yaml; print('PyYAML OK')"
```

정상 예시
- PyYAML OK

---

## 5단계. Dooray API Token 준비
서버가 Dooray 데이터를 읽으려면 API Token이 필요하다.

필요한 것
- DOORAY_API_TOKEN 환경 변수

권장 방법
- Windows 시스템 환경 변수로 등록

설정 방법
1. Windows 검색에서
   - 시스템 환경 변수 편집
   검색 후 실행
2. 아래쪽의 "환경 변수" 버튼 클릭
3. 시스템 변수 또는 사용자 변수에서 새로 만들기
4. 변수 이름 입력
   - DOORAY_API_TOKEN
5. 변수 값 입력
   - 실제 Dooray API Token
6. 확인 눌러 저장
7. PowerShell을 완전히 닫았다가 다시 열기

확인 방법
PowerShell에서 아래 실행

```powershell
$env:DOORAY_API_TOKEN
```

정상이라면
- 토큰 문자열이 출력됨

주의
- 토큰은 메일 본문이나 공유 문서에 쓰지 않는 것이 안전함
- 서버 내부에만 저장 권장

---

## 6단계. PowerShell 실행 정책 확인
일부 서버에서는 .ps1 실행이 막혀 있을 수 있다.

테스트 방법
PowerShell에서 아래 실행

```powershell
Get-ExecutionPolicy
```

만약 너무 엄격해서 스크립트 실행이 막히면,
일단 현재 사용자 기준으로만 완화하는 방법

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

질문이 나오면
- Y 입력 후 Enter

주의
- 회사 보안정책이 있으면 그 기준을 우선 따라야 함

---

## 7단계. 수동 테스트 실행
이 단계가 가장 중요하다.
스케줄러 등록 전에 반드시 사람이 직접 한 번 실행해 봐야 한다.

실행 방법
PowerShell에서 아래 순서대로 입력

```powershell
Set-Location F:\telegram-deadline-bot\telegram-deadline-bot
.\run_latest_preconstruction_validation.ps1
```

이 스크립트가 하는 일
1. Dooray 프로젝트 게시글 최신 목록 조회
2. 각 업무 상세/첨부/로그 조회
3. 규칙에 따라 누락 항목 점검
4. 결과 리포트 저장

정상 실행 시 기대 출력 예시
- saved_report=preconstruction_validation_report_latest.txt
- saved_json=preconstruction_validation_report_latest.json
- validated_count=79

---

## 8단계. 결과 파일 확인
수동 실행이 끝나면 아래 파일이 생성되는지 확인한다.

생성 파일
- F:\telegram-deadline-bot\telegram-deadline-bot\lab_posts_latest.json
- F:\telegram-deadline-bot\telegram-deadline-bot\preconstruction_validation_report_latest.txt
- F:\telegram-deadline-bot\telegram-deadline-bot\preconstruction_validation_report_latest.json

파일 의미
- lab_posts_latest.json
  - Dooray에서 방금 읽어온 최신 업무 원본
- preconstruction_validation_report_latest.txt
  - 사람이 읽기 쉬운 보고서
- preconstruction_validation_report_latest.json
  - 후속 가공/자동 필터링용 데이터

가장 먼저 열어볼 파일
- preconstruction_validation_report_latest.txt

---

## 9단계. 현재 점검 기준 이해하기
이 스크립트는 모든 업무를 자동 처리하는 것이 아니다.
현재 용도는 아래와 같다.

현재 용도
- 프로젝트/업무 보완 점검용

주요 점검 항목
- 공사명
- 공사내용
- 첨부파일
- 계약일
- 착공일
- 준공일

운영 원칙
- 완료된 업무는 수정 대상에서 제외
- 김진광이 포함된 미완료 업무를 우선 점검

즉,
이 환경은 "업무를 대신 처리하는 시스템"이 아니라
"어떤 항목이 비었는지 알려주는 점검 환경"이다.

---

## 10단계. 서버 전원 조건
서버는 아래 상태여도 괜찮다.
- 화면보호기
- 화면 꺼짐
- 잠금 화면

서버는 아래 상태면 안 된다.
- 절전 모드
- 최대 절전 모드
- 네트워크 끊김

즉,
서버는 켜져 있기만 하면 되고,
모니터가 보일 필요는 없다.

---

## 11단계. 나중에 작업 스케줄러 등록
수동 실행이 성공하면 그 다음에 예약 실행을 붙인다.

권장 시각
- 평일 오전 08:30
- 평일 오후 05:00

작업 스케줄러 등록 시 핵심 옵션
- 사용자가 로그온했는지 여부와 관계없이 실행
- 가장 높은 권한으로 실행
- 절전 상태에서는 실행 안 될 수 있으니 절전 비활성화
- 시작 위치(Start in)는 반드시 작업 폴더 지정 권장

권장 시작 위치
- F:\telegram-deadline-bot\telegram-deadline-bot

권장 프로그램/스크립트
- powershell.exe

권장 인수 예시
```powershell
-NoProfile -ExecutionPolicy Bypass -File F:\telegram-deadline-bot\telegram-deadline-bot\run_latest_preconstruction_validation.ps1
```

---

## 12단계. 앞으로 추가할 확장 기능
현재 가능한 것
- 최신 업무 조회
- 보완 필요 항목 리포트 생성

앞으로 붙일 것
- Dooray 메신저 자동 발송

추가 예정 스크립트 예시
- send_included_open_items_to_dooray.ps1

이 스크립트가 하게 될 일
1. latest JSON 리포트 읽기
2. 김진광 포함 업무만 추리기
3. 완료 업무 제외
4. 보완 필요 항목만 요약
5. Dooray 메신저로 direct-send

---

## 초보자 체크리스트
아래 순서대로 체크하면 된다.

### 설치 체크리스트
- [ ] F:\telegram-deadline-bot\telegram-deadline-bot 폴더 생성
- [ ] Python 3.14 설치
- [ ] python --version 확인
- [ ] PyYAML 설치
- [ ] DOORAY_API_TOKEN 환경 변수 등록
- [ ] 필수 파일 3개 복사
- [ ] PowerShell에서 수동 실행 성공
- [ ] TXT/JSON 결과 파일 생성 확인

### 운영 체크리스트
- [ ] 서버 절전 모드 해제
- [ ] 작업 스케줄러 등록
- [ ] 오전/오후 예약 시각 확인
- [ ] 추후 Dooray 메신저 발송 스크립트 추가

---

## 오류가 날 때 가장 먼저 확인할 것
### 1. 토큰 문제
증상
- 인증 실패

확인
```powershell
$env:DOORAY_API_TOKEN
```

### 2. Python 문제
증상
- python not found
- pip not found

확인
```powershell
python --version
C:\Python314\python.exe --version
```

### 3. PyYAML 문제
증상
- ModuleNotFoundError: yaml

해결
```powershell
C:\Python314\python.exe -m pip install pyyaml
```

### 4. 스크립트 실행 정책 문제
증상
- .ps1 실행 차단

확인
```powershell
Get-ExecutionPolicy
```

### 5. 인터넷/방화벽 문제
증상
- Dooray API 호출 실패

확인 포인트
- 서버에서 api.gov-dooray.com 접속 가능한지
- 회사 방화벽 정책 여부

---

## 메일에 넣기 좋은 아주 짧은 요약
- 이 서버는 브라우저 로그인 없이 Dooray API Token만으로 프로젝트 점검 가능
- 초보자 기준으로는 Python 설치, PyYAML 설치, 환경 변수 등록, 스크립트 복사, PowerShell 테스트 실행 순서로 세팅하면 됨
- 현재는 보완 필요 항목 리포트 생성까지 가능하고, 이후 Dooray 메신저 자동발송 스크립트를 붙이면 알림형 운영으로 확장 가능
- 서버는 잠금/화면보호기 상태여도 되지만 절전 모드면 안 됨
