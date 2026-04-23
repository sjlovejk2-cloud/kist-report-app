# 서버용 컴퓨터 설치/운영 정리

## 목적
서버용 컴퓨터에서 평일 정해진 시각에 Dooray 프로젝트를 점검하고,
"김진광이 포함된 미완료 업무" 중 보완이 필요한 항목만 추려서 확인할 수 있게 한다.

현재 기준으로는 아래 두 가지가 이미 가능하다.
1. Dooray 프로젝트 최신 데이터 재조회
2. 보완 필요 항목 자동 점검 리포트 생성

아직 별도 추가 작업이 필요한 부분은 아래이다.
3. 점검 결과를 Dooray 메신저로 자동 발송

즉, 현재는 "점검/보완용 리포트 생성"까지 준비되어 있고,
추후 "Dooray 메신저 자동 발송"을 붙이면 완성형 운영이 된다.

---

## 서버에 둘 작업 폴더
권장 경로 예시
- F:\telegram-deadline-bot\telegram-deadline-bot

이 폴더에 현재 사용 중인 스크립트와 결과 파일을 둔다.

핵심 파일
- validate_preconstruction_meetings.py
  - Dooray 게시글을 읽어 보완 필요 항목을 점검하는 메인 파이썬 스크립트
- preconstruction_validation_rules.yaml
  - 점검 규칙 파일
- run_latest_preconstruction_validation.ps1
  - PowerShell에서 최신 게시글 조회 + 파이썬 점검 실행까지 한 번에 처리하는 실행 스크립트

실행 시 생성되는 결과 파일
- lab_posts_latest.json
  - Dooray에서 최신 조회한 원본 목록
- preconstruction_validation_report_latest.txt
  - 사람이 읽기 쉬운 텍스트 리포트
- preconstruction_validation_report_latest.json
  - 후속 가공용 JSON 리포트

참고 파일
- daily_deadline_report.ps1
  - 예전 Dooray 메신저 direct-send 예제가 들어 있는 스크립트
  - messenger/v1/channels/direct-send 호출 구조 참고 가능

---

## 서버에서 필요한 준비물
1. Windows PowerShell 실행 가능
2. Python 3.14 권장
   - 현재 기준 경로 예시: C:\Python314\python.exe
3. Dooray API Token
4. 인터넷에서 api.gov-dooray.com 접속 가능

중요
- 브라우저에서 Dooray 로그인 상태일 필요는 없음
- API Token만 있으면 됨

---

## 환경 변수
서버에서 반드시 준비할 것
- DOORAY_API_TOKEN

용도
- Dooray API 조회용 인증
- 최신 업무 목록 조회
- 추후 Dooray 메신저 direct-send 발송 시에도 사용 가능

권장 방식
- Windows 시스템 환경 변수 또는 작업 스케줄러 작업 내부에서 설정

주의
- 토큰 문자열은 메일 본문에 직접 쓰지 않는 것이 좋음
- 서버에만 저장하고 공유 문서에는 "환경 변수 설정 필요"로만 적는 것이 안전함

---

## 현재 바로 되는 실행 방법
PowerShell에서 아래 스크립트 실행

```powershell
Set-Location F:\telegram-deadline-bot\telegram-deadline-bot
.\run_latest_preconstruction_validation.ps1
```

이 스크립트가 하는 일
1. Dooray 프로젝트 게시글 최신 목록 조회
2. 각 업무 상세/첨부/로그 조회
3. 규칙에 따라 보완 필요 항목 점검
4. 최신 TXT/JSON 리포트 저장

정상 실행 시 출력 예시
- saved_report=preconstruction_validation_report_latest.txt
- saved_json=preconstruction_validation_report_latest.json
- validated_count=79

---

## 현재 점검 기준
현재는 아래 용도로 사용 중
- 프로젝트/업무 자동 처리용이 아니라
- 프로젝트 보완/정비 점검용

핵심 기준
1. 김진광이 포함된 업무만 의미 있게 본다
2. 완료된 업무는 수정 대상에서 제외한다
3. 미완료 업무에서 아래 누락을 본다
   - 공사명
   - 공사내용
   - 첨부파일
   - 단계별 날짜(계약일/착공일/준공일)

실제 운영 해석
- 자동화 완성품이라기보다 "프로젝트 보완용 점검 도구"

---

## 서버에 설치 후 운영 흐름
### 1단계: 수동 실행 확인
먼저 서버에서 사람이 직접 한 번 실행해 본다.

확인 항목
- 스크립트 오류 없이 종료되는지
- preconstruction_validation_report_latest.txt 생성되는지
- preconstruction_validation_report_latest.json 생성되는지

### 2단계: 작업 스케줄러 등록
권장 시각
- 평일 오전 08:30
- 평일 오후 05:00

이 단계에서 필요한 조건
- 서버 컴퓨터 전원 켜짐
- 절전 모드 비활성화
- 로그온 여부와 관계없이 작업 실행 설정

### 3단계: 결과 확인 방식 결정
선택지 A. 서버 폴더 리포트만 확인
- 가장 단순함

선택지 B. Dooray 메신저로 자동 발송
- 가장 편함
- 다만 발송용 래퍼 스크립트 추가 필요

---

## 서버 전원/운영 조건
가능
- 화면보호기
- 화면 꺼짐
- 잠금 화면

불가
- 절전 모드
- 최대 절전 모드
- 네트워크 끊김

즉,
서버는 켜져 있기만 하면 되고 화면이 켜져 있을 필요는 없다.

---

## 앞으로 추가할 스크립트(권장)
현재 없는 것
- 점검 결과를 "김진광 포함 미완료 업무만 추려서 Dooray 메신저로 발송"하는 전용 스크립트

추가 권장 파일 예시
- send_included_open_items_to_dooray.ps1
  - latest JSON 리포트를 읽음
  - assignee/watcher 중 김진광 포함 업무만 추림
  - 완료 업무 제외
  - 보완 필요 항목만 짧게 메시지로 만듦
  - Dooray messenger direct-send 호출

이 파일까지 붙으면 최종 구조는 아래처럼 됨
1. run_latest_preconstruction_validation.ps1
   - 최신 점검 실행
2. send_included_open_items_to_dooray.ps1
   - 결과 요약 후 Dooray 메신저 발송

---

## 추천 최종 운영 구조
### 현재 즉시 가능한 구조
- 서버에 작업 폴더 복사
- Python/토큰 세팅
- run_latest_preconstruction_validation.ps1 수동 실행
- 결과 파일 확인

### 다음 단계 권장 구조
- 작업 스케줄러 2회 등록
- 오전 08:30 점검
- 오후 05:00 점검
- 이후 Dooray 메신저 자동 발송 래퍼 추가

---

## 메일로 전달할 때 핵심 요약 문안
- 서버용 컴퓨터에 Dooray 점검 스크립트를 설치하면 브라우저 로그인 없이 API Token 기반으로 프로젝트 점검 가능
- 현재는 최신 업무 재조회 및 보완 필요 항목 리포트 생성까지 가능
- 추후 Dooray 메신저 direct-send 스크립트를 추가하면 평일 08:30 / 17:00 자동 알림 구조로 확장 가능
- 서버는 화면보호기/잠금 상태여도 되지만 절전 모드는 비활성화해야 함
- 완료 업무는 제외하고, 김진광 포함 미완료 업무만 보완 대상으로 운영
