# GL Input Copy Tool

Python `tkinter` GUI로 Review 템플릿 xlsx와 ERP Export xlsx를 선택한 뒤, Export 첫 번째 시트의 데이터를 Review 템플릿의 `2. GL Input` 시트에 복사해 새 xlsx로 저장합니다.

## 실행

exe로 실행:

```powershell
.\dist\GLInputCopyTool.exe
```

Python으로 실행:

```powershell
py -m pip install -r requirements.txt
py .\gl_input_copy_gui.py
```

## 복사 기준

- Review 템플릿에서는 이름이 정확히 `2. GL Input`인 시트를 사용합니다.
- ERP Export에서는 첫 번째 시트를 사용합니다.
- 데이터는 행/열 위치가 아니라 아래 헤더명 기준으로 복사합니다.
  - `날짜`
  - `계정코드`
  - `차변(EUR)`
  - `대변(EUR)`
  - `거래처명`
  - `적요`
- 양쪽 시트의 상위 30행 안에서 위 헤더들이 모두 있는 행을 자동으로 찾습니다.
- Review의 기존 입력 데이터는 위 헤더 컬럼 범위에서 삭제한 뒤 새 데이터로 채웁니다.
- GL Input 입력 후, 입력된 데이터 마지막 행보다 아래에 있는 행은 삭제합니다.
- Review의 수식 컬럼은 삭제하거나 덮어쓰지 않습니다.

## 진행 상황 표시

실행 버튼을 누르면 파일 열기, 헤더 탐색, 기존 데이터 삭제, 데이터 입력, 하단 행 삭제, 저장 단계가 progress bar와 상태 문구로 표시됩니다.

## 샘플 파일

`samples` 폴더에 테스트용 파일이 있습니다.

- `Review_Template_Sample.xlsx`: Review 템플릿 예시
- `ERP_Export_Sample.xlsx`: ERP Export 예시. 컬럼 순서가 Review와 다르게 되어 있습니다.
- `Expected_Result_Sample.xlsx`: 위 두 파일을 실행했을 때 기대되는 결과 예시
