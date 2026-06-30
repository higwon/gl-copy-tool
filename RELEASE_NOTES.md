# Release Notes

## v0.2.0

첫 배포용 릴리즈입니다.

### 추가

- 앱 우측 상단에 버전과 제작자 표시를 추가했습니다.
- `Import Data 양식` 콤보박스를 추가했습니다.
- 등록된 양식:
  - `Korea - Standard`
  - `Germany - SAP Export`
  - `Japan - Standard`
  - `China - Standard`
  - `Vietnam - Standard`
- xlsx 파일 경로 입력칸에 드래그앤드랍을 지원합니다.
- 결과 표시 영역에 스크롤 가능한 읽기 전용 뷰어를 추가했습니다.
- GitHub Actions 빌드 워크플로를 추가했습니다.
- 태그 push 시 GitHub Release를 만들고 exe를 첨부하도록 구성했습니다.

### 변경

- 버전을 `v0.2.0`으로 올렸습니다.
- `계정코드`는 Import Data에서 숫자로 들어와도 결과 파일에는 텍스트로 저장되도록 변경했습니다.
- 로컬 exe는 계속 빌드해 테스트하되, Git에서는 exe 추적을 제외하고 GitHub Actions artifact/Release asset으로 관리하도록 정리했습니다.

### 검증

- Python 문법 체크를 통과했습니다.
- 샘플 파일 기반 복사 함수 실행을 확인했습니다.
- Windows용 `GLInputCopyTool.exe` 로컬 빌드를 확인했습니다.
