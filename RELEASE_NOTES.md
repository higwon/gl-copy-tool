# Release Notes

## v0.4.1

- `Slovenia` 국가 양식을 추가했습니다.
- `Germany` 국가 양식을 추가했습니다.
- GL Auto 템플릿의 `차변(...)`, `대변(...)` 헤더는 괄호 안 통화 단위가 달라도 인식하도록 개선했습니다.

## v0.4.0

- 국가별 파서를 `countries` 패키지로 분리하고 registry 기반 확장 구조로 변경했습니다.
- 체코와 슬로바키아의 MD/DAL 변환을 공통 복식부기 엔진으로 통합했습니다.
- 입력, 진행 상태, 실행 결과 중심으로 UI 레이아웃과 색상 체계를 정리했습니다.
- 대용량 입력 진행 갱신 횟수를 줄여 UI 응답성을 개선했습니다.
- 결과 파일이 Source GL을 덮어쓰지 못하도록 경로 검증을 추가했습니다.
- 7개 국가 파서 자동 회귀 테스트를 추가했습니다.

## v0.3.5

- `Slovakia` 국가 양식을 추가했습니다.
- Slovakia GL의 MD/DAL 복식부기 행을 계정별 차변/대변 2행으로 변환하고 EUR 금액을 입력합니다.
- 국가 선택 콤보박스에 한글과 영문 국가명을 함께 표시합니다.

## v0.3.4

- `Czech Republic` 국가 양식을 추가했습니다.
- Czech GL의 MD/DAL 복식부기 행을 계정별 차변/대변 2행으로 변환하고 CZK 금액을 입력합니다.
- 국가 선택 옆의 지원 국가 안내 문구를 제거했습니다.

## v0.3.3

- `Poland` 국가 양식을 추가했습니다.
- Poland `Zapisy księgowe`의 원본 헤더 기준 매핑과 PLN 차변/대변 입력을 지원합니다.
- Hungary Novitax export는 2행 한글 힌트 없이 원본 헤더 다음 행부터 처리하도록 기준을 확정했습니다.
- 대용량 데이터 행 확장 시 첫 데이터 행의 입력 서식과 날짜 형식을 유지하도록 수정했습니다.

## v0.3.2

- `Hungary` 국가 양식을 추가했습니다.
- Hungary Novitax export의 원본 헤더 기준 매핑을 지원합니다.
- 국가 선택 콤보박스 값을 `Korea`, `Netherlands`, `Austria`, `Hungary`처럼 나라명만 표시하도록 정리했습니다.
- README의 버전, 실행 파일명, 지원 국가 설명을 갱신했습니다.
- 릴리즈 노트를 버전별 변경사항만 남기는 간결한 형식으로 정리했습니다.

## v0.3.1

- 분리된 차변/대변 Source GL 금액을 GL Input 입력 전 양수로 정규화했습니다.
- GL Input 계정코드 타입을 GL Auto 템플릿 `1. COA` 시트의 실제 값 타입에 맞추도록 변경했습니다.
- UI 용어를 Source GL / 국가 선택 기준으로 변경했습니다.
- PyInstaller 실행 파일명에 앱 버전을 붙이도록 변경했습니다.
- GitHub Actions Release asset에서 SHA256 파일 생성을 제거했습니다.

## v0.3.0

- Netherlands GL Transactions Report 양식을 추가했습니다.
- Austria FIBU Export 양식을 추가했습니다.
- 지원하는 국가/Source GL 양식만 콤보박스에 표시하도록 정리했습니다.
- 대용량 파일 처리 중 진행 표시와 경과 시간 표시를 개선했습니다.
- PyInstaller 빌드에 `tkinterdnd2` 리소스를 포함했습니다.
- Source GL 행 수가 기존 GL Input 표보다 많을 때 표 범위와 행 서식이 확장되도록 수정했습니다.

## v0.2.2

- 앱 헤더, 상태 영역, 완료/오류 팝업에 이미지 리소스를 추가했습니다.
- Windows 실행 파일 아이콘을 추가했습니다.
- PyInstaller 산출물에 `assets` 폴더가 포함되도록 설정했습니다.

## v0.2.1

- 배포 방식을 단일 exe에서 폴더형 zip으로 변경했습니다.
- GitHub Actions가 zip 릴리즈 파일을 생성하도록 변경했습니다.
- README를 폴더형 zip 사용 방식에 맞게 갱신했습니다.

## v0.2.0

- Tkinter 기반 초기 GUI를 추가했습니다.
- GL Auto 템플릿, Source GL, 결과 저장 경로 선택 기능을 추가했습니다.
- Source GL 양식 선택, 파일 드래그앤드랍, 실행 결과 요약, GitHub Actions 빌드/릴리즈 흐름을 추가했습니다.
- 로컬 exe 빌드 산출물을 Git 추적에서 제외했습니다.
