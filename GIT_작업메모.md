# Git 작업 메모

프로젝트 작업 폴더
- `F:\kist-report-app\project`

일반 작업 순서
```powershell
cd F:\kist-report-app\project
git status
git add .
git commit -m "작업 내용"
git push
```

간단 실행 방법
- `git_push_update.bat` 실행
- 커밋 메시지 입력
- 자동으로 `git add .` → `git commit` → `git push` 진행

서버 실행
- 바탕화면의 `서버실행.bat` 실행
- 실제 앱 실행 경로는 `F:\kist-report-app\project` 기준

문제 발생 시 확인
1. `git status` 로 변경 파일 확인
2. `git remote -v` 로 원격 저장소 주소 확인
3. `git push` 실패 시 인증 상태 확인
4. 서버 실행 실패 시 Python/Streamlit 설치 여부 확인

원격 저장소
- `https://github.com/sjlovejk2-cloud/kist-report-app.git`
