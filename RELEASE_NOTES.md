# Release Notes

## v0.3.1

GL Input 입력값 정규화와 배포 파일명을 정리한 패치 릴리즈입니다.

### 변경

- 차변/대변이 분리된 Source GL 양식은 `차변(EUR)`, `대변(EUR)` 값을 모두 양수 금액으로 변환해 입력합니다.
- `계정코드`는 GL Auto 템플릿의 `1. COA` 시트에 저장된 실제 `계정코드` 타입을 따라 텍스트 또는 숫자로 입력합니다.
- UI 라벨을 `Import Data 양식`에서 `국가 선택`으로, `Import Data xlsx`에서 `Source GL xlsx`로 변경했습니다.
- PyInstaller 실행 파일명에 앱 버전을 붙이도록 변경했습니다. 예: `GLInputCopyTool-v0.3.1.exe`
- GitHub Actions Release asset에서 SHA256 파일 생성을 제거하고 zip 파일만 배포합니다.

### 검증

- Python 문법 체크를 통과했습니다.
- Austria 실제 GL 파일 변환에서 차변/대변 금액이 양수로 입력되는 것을 확인했습니다.

## v0.3.0

국가별 Import Data 양식 지원과 대용량 처리 진행 표시를 개선한 릴리즈입니다.

### 추가

- `Netherlands - GL Transactions Report` 양식을 추가했습니다.
  - 계정 블록 행에서 GL 계정코드를 추출합니다.
  - 날짜가 있는 거래행만 GL Input으로 변환합니다.
  - `Date`, `Debit`, `Credit`, `Account name`, `Description` 컬럼을 표준 GL Input 컬럼으로 매핑합니다.
- `Austria - FIBU Export` 양식을 추가했습니다.
  - `Beleg-Dat`, `Kto-Nr`, `GW-Soll`, `GW-Haben`, `Text` 컬럼을 표준 GL Input 컬럼으로 매핑합니다.
  - `GW-Haben`은 대변 금액으로 양수 처리합니다.
- Import Data 양식 콤보박스를 실제 지원 양식만 표시하도록 정리했습니다.
- 대용량 파일 처리 중 progress bar가 멈춘 것처럼 보이지 않도록 개선했습니다.
  - 파일 열기/저장 중에는 움직이는 progress bar를 표시합니다.
  - 작업 중 경과 시간을 표시합니다.
  - 거래행 파싱 중 처리 행 수를 주기적으로 표시합니다.
- PyInstaller 빌드에 `tkinterdnd2` 리소스를 포함해 드래그앤드랍 동작을 보강했습니다.

### 수정

- Import Data 행 수가 GL Input 기존 표보다 많을 때 Excel 표 범위와 행 서식이 확장되지 않던 문제를 수정했습니다.
- 확장된 행에서도 날짜 형식, 계정코드 텍스트 형식, 수식 컬럼이 유지되도록 수정했습니다.

### 검증

- Python 문법 체크를 통과했습니다.
- Korea 샘플 변환을 확인했습니다.
- Netherlands 실제 GL 파일 17,360행 변환을 확인했습니다.
- Austria 실제 GL 파일 7,138행 변환을 확인했습니다.
- 로컬 Windows 폴더형 PyInstaller 빌드를 확인했습니다.

## v0.2.2

앱 이미지와 아이콘을 적용한 UI 개선 릴리즈입니다.

### 추가

- 앱 상단 타이틀 왼쪽에 로고 이미지를 추가했습니다.
- Windows 실행 파일 아이콘을 커스텀 이미지로 변경했습니다.
- 하단 결과 영역 왼쪽에 성공/실패 상태 이미지를 추가했습니다.
- 성공/실패 알림을 이미지가 포함된 커스텀 팝업으로 변경했습니다.
- PyInstaller 빌드 산출물에 `assets` 폴더가 포함되도록 설정했습니다.

### 검증

- Python 문법 체크를 통과했습니다.
- Tkinter 앱 생성 시 모든 이미지 assets가 로드되는 것을 확인했습니다.

## v0.2.1

단일 exe 배포에서 폴더형 zip 배포로 전환한 패치 릴리즈입니다.

### 변경

- PyInstaller 빌드 방식을 단일 파일 exe에서 폴더형 `onedir` 빌드로 변경했습니다.
- GitHub Actions가 `GLInputCopyTool-v0.2.1.zip`을 생성하도록 변경했습니다.
- Release에는 zip 파일과 SHA256 파일을 함께 첨부하도록 변경했습니다.
- 앱 버전을 `v0.2.1`로 올렸습니다.
- README를 zip 배포 방식 기준으로 갱신했습니다.

### 이유

- 서명되지 않은 PyInstaller 단일 exe가 Windows Defender 또는 브라우저에서 바이러스로 탐지될 수 있어, 오탐 가능성을 줄이기 위해 폴더형 zip 배포로 변경했습니다.
- 코드 서명은 아직 적용하지 않습니다.

### 사용 방법

1. `GLInputCopyTool-v0.2.1.zip`을 다운로드합니다.
2. zip 압축을 풉니다.
3. 압축을 푼 폴더 안의 `GLInputCopyTool.exe`를 실행합니다.

### 검증

- Python 문법 체크를 통과했습니다.
- PyInstaller 폴더형 로컬 빌드를 확인했습니다.
- 로컬 zip 패키지 생성을 확인했습니다.

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
