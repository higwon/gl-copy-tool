# Release Notes

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
