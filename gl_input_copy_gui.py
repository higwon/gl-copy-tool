import os
import tkinter as tk
from tkinter import filedialog, messagebox


TARGET_SHEET_NAME = "2. GL Input"
MATCH_HEADERS = ["날짜", "계정코드", "차변(EUR)", "대변(EUR)", "거래처명", "적요"]
HEADER_SCAN_ROWS = 30


def is_formula_cell(cell):
    value = cell.value
    return cell.data_type == "f" or (isinstance(value, str) and value.startswith("="))


def is_read_only_merged_cell(cell):
    return cell.__class__.__name__ == "MergedCell"


def normalize_header(value):
    if value is None:
        return ""
    return "".join(str(value).split()).lower()


def find_header_row_and_columns(sheet, required_headers):
    required = {normalize_header(header): header for header in required_headers}

    for row in sheet.iter_rows(max_row=min(sheet.max_row, HEADER_SCAN_ROWS)):
        found = {}
        for cell in row:
            normalized = normalize_header(cell.value)
            if normalized in required and normalized not in found:
                found[normalized] = cell.column

        if all(normalize_header(header) in found for header in required_headers):
            return row[0].row, {
                required[normalized]: column for normalized, column in found.items()
            }

    missing_text = ", ".join(required_headers)
    raise ValueError(
        f'"{sheet.title}" 시트의 상위 {HEADER_SCAN_ROWS}행에서 '
        f"필수 헤더를 찾지 못했습니다: {missing_text}"
    )


def row_has_data(sheet, row_number, columns):
    for column in columns:
        if sheet.cell(row=row_number, column=column).value not in (None, ""):
            return True
    return False


def find_formula_columns(sheet):
    formula_columns = set()
    for row in sheet.iter_rows():
        for cell in row:
            if is_read_only_merged_cell(cell):
                continue
            if is_formula_cell(cell):
                formula_columns.add(cell.column)
    return formula_columns


def clear_target_data(sheet, header_row, target_columns, formula_columns, min_clear_row):
    clear_until = max(sheet.max_row, min_clear_row)
    for row in sheet.iter_rows(min_row=header_row + 1, max_row=clear_until):
        for cell in row:
            if is_read_only_merged_cell(cell):
                continue
            if cell.column in target_columns and cell.column not in formula_columns:
                cell.value = None


def copy_cell_value_only(source_cell, target_cell):
    target_cell.value = source_cell.value


def copy_export_to_gl_input(template_path, export_path, output_path):
    try:
        from openpyxl import load_workbook
    except ModuleNotFoundError as exc:
        raise ValueError(
            "openpyxl이 설치되어 있지 않습니다. PowerShell에서 "
            "'py -m pip install -r requirements.txt'를 실행한 뒤 다시 시도하세요."
        ) from exc

    template_wb = load_workbook(template_path)
    export_wb = load_workbook(export_path, data_only=False)

    if TARGET_SHEET_NAME not in template_wb.sheetnames:
        raise ValueError(f'Review template에 "{TARGET_SHEET_NAME}" 시트가 없습니다.')

    target_ws = template_wb[TARGET_SHEET_NAME]
    source_ws = export_wb.worksheets[0]

    formula_columns = find_formula_columns(target_ws)
    source_header_row, source_columns = find_header_row_and_columns(source_ws, MATCH_HEADERS)
    target_header_row, target_columns = find_header_row_and_columns(target_ws, MATCH_HEADERS)
    source_data_rows = [
        row_number
        for row_number in range(source_header_row + 1, source_ws.max_row + 1)
        if row_has_data(source_ws, row_number, source_columns.values())
    ]

    clear_target_data(
        target_ws,
        target_header_row,
        set(target_columns.values()),
        formula_columns,
        target_header_row + len(source_data_rows),
    )

    for row_offset, source_row_number in enumerate(source_data_rows, start=1):
        target_row_number = target_header_row + row_offset
        for header in MATCH_HEADERS:
            target_column = target_columns[header]
            if target_column in formula_columns:
                continue

            source_cell = source_ws.cell(
                row=source_row_number,
                column=source_columns[header],
            )
            target_cell = target_ws.cell(row=target_row_number, column=target_column)
            if is_read_only_merged_cell(target_cell):
                continue
            if is_formula_cell(target_cell):
                continue

            copy_cell_value_only(source_cell, target_cell)

    template_wb.save(output_path)


class GlInputCopyApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("GL Input Copy Tool")
        self.geometry("720x250")
        self.resizable(False, False)

        self.template_path = tk.StringVar()
        self.export_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        container = tk.Frame(self, padx=16, pady=16)
        container.pack(fill=tk.BOTH, expand=True)

        self._add_file_row(
            container,
            row=0,
            label="Review 템플릿 xlsx",
            variable=self.template_path,
            command=self.select_template,
        )
        self._add_file_row(
            container,
            row=1,
            label="ERP Export xlsx",
            variable=self.export_path,
            command=self.select_export,
        )
        self._add_file_row(
            container,
            row=2,
            label="결과 저장 경로",
            variable=self.output_path,
            command=self.select_output,
        )

        run_button = tk.Button(container, text="실행", command=self.run, width=16, height=2)
        run_button.grid(row=3, column=2, sticky="e", pady=(20, 0))

        container.grid_columnconfigure(1, weight=1)

    def _add_file_row(self, parent, row, label, variable, command):
        tk.Label(parent, text=label, anchor="w", width=18).grid(
            row=row, column=0, sticky="w", pady=8
        )
        tk.Entry(parent, textvariable=variable, width=70).grid(
            row=row, column=1, sticky="ew", padx=(8, 8), pady=8
        )
        tk.Button(parent, text="찾기", command=command, width=10).grid(
            row=row, column=2, sticky="e", pady=8
        )

    def select_template(self):
        path = filedialog.askopenfilename(
            title="Review 템플릿 xlsx 선택",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.template_path.set(path)
            self._suggest_output_path()

    def select_export(self):
        path = filedialog.askopenfilename(
            title="ERP Export xlsx 선택",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.export_path.set(path)
            self._suggest_output_path()

    def select_output(self):
        path = filedialog.asksaveasfilename(
            title="결과 저장 경로 선택",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.output_path.set(path)

    def _suggest_output_path(self):
        if self.output_path.get() or not self.template_path.get():
            return

        base_dir = os.path.dirname(self.template_path.get())
        base_name = os.path.splitext(os.path.basename(self.template_path.get()))[0]
        self.output_path.set(os.path.join(base_dir, f"{base_name}_result.xlsx"))

    def validate_inputs(self):
        template_path = self.template_path.get().strip()
        export_path = self.export_path.get().strip()
        output_path = self.output_path.get().strip()

        if not template_path:
            raise ValueError("Review 템플릿 xlsx를 선택하세요.")
        if not export_path:
            raise ValueError("ERP Export xlsx를 선택하세요.")
        if not output_path:
            raise ValueError("결과 저장 경로를 선택하세요.")
        if not os.path.isfile(template_path):
            raise ValueError("Review 템플릿 파일을 찾을 수 없습니다.")
        if not os.path.isfile(export_path):
            raise ValueError("ERP Export 파일을 찾을 수 없습니다.")
        if os.path.abspath(template_path) == os.path.abspath(output_path):
            raise ValueError("결과 파일은 Review 템플릿과 다른 경로로 저장하세요.")

        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.isdir(output_dir):
            raise ValueError("결과 저장 폴더를 찾을 수 없습니다.")

        return template_path, export_path, output_path

    def run(self):
        try:
            template_path, export_path, output_path = self.validate_inputs()
            copy_export_to_gl_input(template_path, export_path, output_path)
            messagebox.showinfo("완료", f"결과 파일을 저장했습니다.\n\n{output_path}")
        except Exception as exc:
            messagebox.showerror("오류", str(exc))


if __name__ == "__main__":
    app = GlInputCopyApp()
    app.mainloop()
