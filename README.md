# GL Input Copy Tool

GL Auto 템플릿 xlsx와 Source GL xlsx를 선택하면, Source GL의 첫 번째 시트 데이터를 GL Auto 템플릿의 `2. GL Input` 시트로 복사해 결과 xlsx를 저장하는 Windows GUI 도구입니다.

현재 버전: `v0.3.4`

## 다운로드

최신 버전은 GitHub Releases에서 받습니다.

[최신 릴리즈 다운로드](https://github.com/higwon/gl-copy-tool/releases/latest)

`v0.2.1`부터는 백신 오탐 가능성을 줄이기 위해 단일 exe 대신 폴더형 zip으로 배포합니다.

사용 방법:

1. Release에서 `GLInputCopyTool-v0.3.4.zip`을 다운로드합니다.
2. zip 압축을 풉니다.
3. 압축을 푼 폴더 안의 `GLInputCopyTool-v0.3.4.exe`를 실행합니다.

참고:

- 코드 서명은 아직 적용하지 않았습니다.
- Windows 또는 브라우저가 경고를 표시할 수 있습니다.

## 로컬 실행

Python으로 실행:

```powershell
python -m pip install -r requirements.txt
python .\gl_input_copy_gui.py
```

로컬 폴더형 빌드:

```powershell
python -m pip install -r requirements.txt
python -m pip install pyinstaller
python -m PyInstaller GLInputCopyTool.spec --noconfirm
```

빌드 후 실행:

```powershell
.\dist\GLInputCopyTool\GLInputCopyTool-v0.3.4.exe
```

로컬 테스트용 빌드는 계속 만들어도 됩니다. 배포용 파일은 Git에 직접 커밋하지 않고 GitHub Actions artifact 또는 Release asset으로 관리합니다.

## 주요 기능

- GL Auto 템플릿의 `2. GL Input` 시트에 Source GL을 복사합니다.
- Source GL은 첫 번째 시트를 사용합니다.
- 국가 선택 콤보박스로 Source GL 양식을 선택할 수 있습니다.
- Korea, Netherlands, Austria, Hungary, Poland, Czech Republic 국가 양식을 지원합니다.
- 대용량 파일 처리 중 progress bar와 경과 시간이 계속 갱신됩니다.
- xlsx 파일 경로 입력칸에 파일을 드래그앤드랍할 수 있습니다.
- 실행 결과 영역은 긴 메시지와 파일 경로를 볼 수 있도록 스크롤을 지원합니다.
- 계정코드는 GL Auto 템플릿의 `1. COA` 시트에 저장된 `계정코드` 타입에 맞춰 텍스트 또는 숫자로 저장합니다.
- GL Auto 템플릿의 수식 컬럼은 삭제하거나 덮어쓰지 않고, 입력 행 수에 맞춰 채웁니다.

## 복사 기준

GL Auto 템플릿에는 이름이 정확히 `2. GL Input`인 시트가 있어야 합니다.

기본 복사 대상 헤더:

- `날짜`
- `계정코드`
- `차변(EUR)`
- `대변(EUR)`
- `거래처명`
- `적요`

동작 방식:

- Source GL과 GL Auto 템플릿 모두 상위 30행 안에서 필수 헤더를 찾습니다.
- 컬럼 위치가 달라도 헤더명 기준으로 매칭합니다.
- GL Auto 템플릿의 기존 입력 데이터는 대상 헤더 컬럼 범위에서 삭제한 뒤 새 데이터로 채웁니다.
- 입력 데이터 마지막 행 아래의 남는 행은 삭제합니다.
- Excel 표가 있는 경우 입력 데이터 마지막 행까지만 포함하도록 표 범위를 조정합니다.

## 국가별 Source GL 양식

현재 UI에 등록된 국가 양식:

- `Korea`
- `Netherlands`
- `Austria`
- `Hungary`
- `Poland`
- `Czech Republic`

국가별 Source GL 구조에 맞춰 컬럼 매핑, 금액 처리, 거래처명/적요 매핑 방식을 변환합니다.

## 샘플 파일

`samples` 폴더에는 테스트용 파일이 있습니다.

- `Review_Template_Sample.xlsx`: GL Auto 템플릿 예시
- `ERP_Export_Sample.xlsx`: Source GL 예시
- `Expected_Result_Sample.xlsx`: 샘플 실행 시 기대하는 결과 예시

## 협업 및 변경 추적

이 저장소는 두 PC에서 같은 GitHub repo를 기준으로 작업합니다. 작업 전후에는 아래 흐름을 기준으로 변경사항을 확인합니다.

작업 시작 전:

```powershell
git status
git pull origin main
```

작업 완료 후:

```powershell
git status
git diff --stat
git add <변경 파일>
git commit -m "변경 내용 요약"
git push origin main
```

다른 PC에서 이어서 작업할 때:

```powershell
git pull origin main
```

주의:

- `samples/ERP_Export_Sample.xlsx`처럼 테스트 중 임시로 바뀐 파일은 커밋 전에 포함 여부를 확인합니다.
- 로컬 테스트용 빌드 산출물은 Git에 직접 커밋하지 않습니다.
- 배포용 zip은 GitHub Actions artifact 또는 Release asset으로 관리합니다.

## GitHub Actions와 릴리즈

일반 push 또는 PR:

- Python 문법 체크
- PyInstaller 폴더형 빌드 확인
- `GLInputCopyTool-<ref>.zip`을 Actions artifact로 업로드

버전 태그 push:

```powershell
git tag v0.3.4
git push origin v0.3.4
```

태그가 `v*` 형식이면 GitHub Actions가 Release를 만들고 zip 파일을 Release asset으로 첨부합니다.

## 릴리즈 노트

릴리즈 변경사항은 [RELEASE_NOTES.md](./RELEASE_NOTES.md)에 기록합니다.
