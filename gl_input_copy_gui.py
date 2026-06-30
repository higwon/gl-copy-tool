import copy
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


TARGET_SHEET_NAME = "2. GL Input"
MATCH_HEADERS = ["날짜", "계정코드", "차변(EUR)", "대변(EUR)", "거래처명", "적요"]
HEADER_SCAN_ROWS = 30


def report_progress(progress_callback, percent, message):
    if progress_callback:
        progress_callback(percent, message)


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


def delete_rows_below_input(sheet, header_row, data_row_count):
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.cell import range_boundaries

    first_delete_row = header_row + data_row_count + 1
    last_keep_row = first_delete_row - 1

    for table in sheet.tables.values():
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        if min_row <= last_keep_row < max_row:
            table.ref = (
                f"{get_column_letter(min_col)}{min_row}:"
                f"{get_column_letter(max_col)}{last_keep_row}"
            )
            if table.autoFilter:
                table.autoFilter.ref = table.ref

    if sheet.max_row >= first_delete_row:
        sheet.delete_rows(first_delete_row, sheet.max_row - first_delete_row + 1)

    for row_number in list(sheet.row_dimensions):
        if row_number > last_keep_row:
            del sheet.row_dimensions[row_number]


def fill_formula_columns(sheet, header_row, data_row_count, formula_columns):
    if data_row_count <= 0:
        return

    from openpyxl.formula.translate import Translator

    first_data_row = header_row + 1
    last_data_row = header_row + data_row_count

    for column in sorted(formula_columns):
        template_cell = None
        for row_number in range(first_data_row, sheet.max_row + 1):
            candidate = sheet.cell(row=row_number, column=column)
            if is_read_only_merged_cell(candidate):
                continue
            if is_formula_cell(candidate):
                template_cell = candidate
                break

        if not template_cell:
            continue

        for row_number in range(first_data_row, last_data_row + 1):
            target_cell = sheet.cell(row=row_number, column=column)
            if is_read_only_merged_cell(target_cell):
                continue

            try:
                target_cell.value = Translator(
                    template_cell.value,
                    origin=template_cell.coordinate,
                ).translate_formula(target_cell.coordinate)
            except Exception:
                target_cell.value = template_cell.value

            if template_cell.has_style:
                target_cell._style = copy.copy(template_cell._style)
            if template_cell.number_format:
                target_cell.number_format = template_cell.number_format


def copy_cell_value_only(source_cell, target_cell):
    target_cell.value = source_cell.value


def copy_export_to_gl_input(template_path, export_path, output_path, progress_callback=None):
    try:
        from openpyxl import load_workbook
    except ModuleNotFoundError as exc:
        raise ValueError(
            "openpyxl이 설치되어 있지 않습니다. PowerShell에서 "
            "'py -m pip install -r requirements.txt'를 실행한 뒤 다시 시도하세요."
        ) from exc

    report_progress(progress_callback, 5, "파일을 여는 중...")
    template_wb = load_workbook(template_path)
    export_wb = load_workbook(export_path, data_only=False)

    report_progress(progress_callback, 15, "Review 템플릿 시트를 확인하는 중...")
    if TARGET_SHEET_NAME not in template_wb.sheetnames:
        raise ValueError(f'Review template에 "{TARGET_SHEET_NAME}" 시트가 없습니다.')

    target_ws = template_wb[TARGET_SHEET_NAME]
    source_ws = export_wb.worksheets[0]

    report_progress(progress_callback, 25, "헤더를 찾는 중...")
    formula_columns = find_formula_columns(target_ws)
    source_header_row, source_columns = find_header_row_and_columns(source_ws, MATCH_HEADERS)
    target_header_row, target_columns = find_header_row_and_columns(target_ws, MATCH_HEADERS)

    report_progress(progress_callback, 35, "ERP Export 데이터를 읽는 중...")
    source_data_rows = [
        row_number
        for row_number in range(source_header_row + 1, source_ws.max_row + 1)
        if row_has_data(source_ws, row_number, source_columns.values())
    ]

    report_progress(progress_callback, 45, "기존 GL Input 데이터를 삭제하는 중...")
    clear_target_data(
        target_ws,
        target_header_row,
        set(target_columns.values()),
        formula_columns,
        target_header_row + len(source_data_rows),
    )

    total_rows = max(len(source_data_rows), 1)
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

        percent = 45 + int((row_offset / total_rows) * 35)
        report_progress(
            progress_callback,
            percent,
            f"GL Input에 데이터 입력 중... ({row_offset}/{len(source_data_rows)})",
        )

    report_progress(progress_callback, 82, "수식 컬럼을 정리하는 중...")
    fill_formula_columns(target_ws, target_header_row, len(source_data_rows), formula_columns)

    report_progress(progress_callback, 88, "입력 데이터 아래 행을 삭제하는 중...")
    delete_rows_below_input(target_ws, target_header_row, len(source_data_rows))

    report_progress(progress_callback, 95, "결과 파일을 저장하는 중...")
    template_wb.save(output_path)
    report_progress(progress_callback, 100, "완료")


class GlInputCopyApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("GL Input Copy Tool")
        self.geometry("780x360")
        self.resizable(False, False)
        self.configure(bg="#F5F7FA")

        self.template_path = tk.StringVar()
        self.export_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status_text = tk.StringVar(value="대기 중")
        self.progress_value = tk.IntVar(value=0)

        self.run_button = None
        self.style = ttk.Style(self)
        self._configure_style()
        self._build_ui()

    def _configure_style(self):
        self.style.theme_use("clam")
        self.style.configure("App.TFrame", background="#F5F7FA")
        self.style.configure("Card.TFrame", background="#FFFFFF", relief="flat")
        self.style.configure(
            "Title.TLabel",
            background="#F5F7FA",
            foreground="#1F2937",
            font=("Segoe UI", 16, "bold"),
        )
        self.style.configure(
            "Subtitle.TLabel",
            background="#F5F7FA",
            foreground="#6B7280",
            font=("Segoe UI", 9),
        )
        self.style.configure(
            "Field.TLabel",
            background="#FFFFFF",
            foreground="#374151",
            font=("Segoe UI", 9, "bold"),
        )
        self.style.configure(
            "Status.TLabel",
            background="#FFFFFF",
            foreground="#4B5563",
            font=("Segoe UI", 9),
        )
        self.style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(18, 10),
            background="#2563EB",
            foreground="#FFFFFF",
        )
        self.style.map(
            "Primary.TButton",
            background=[("active", "#1D4ED8"), ("disabled", "#9CA3AF")],
            foreground=[("disabled", "#F3F4F6")],
        )
        self.style.configure("Browse.TButton", font=("Segoe UI", 9), padding=(10, 6))
        self.style.configure(
            "Blue.Horizontal.TProgressbar",
            troughcolor="#E5E7EB",
            background="#2563EB",
            bordercolor="#E5E7EB",
            lightcolor="#2563EB",
            darkcolor="#2563EB",
        )

    def _build_ui(self):
        container = ttk.Frame(self, padding=20, style="App.TFrame")
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(container, text="GL Input Copy Tool", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            container,
            text="ERP Export 데이터를 Review 템플릿의 2. GL Input 시트로 옮깁니다.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 16))

        card = ttk.Frame(container, padding=18, style="Card.TFrame")
        card.grid(row=2, column=0, sticky="ew")

        self._add_file_row(
            card,
            row=0,
            label="Review 템플릿 xlsx",
            variable=self.template_path,
            command=self.select_template,
        )
        self._add_file_row(
            card,
            row=1,
            label="ERP Export xlsx",
            variable=self.export_path,
            command=self.select_export,
        )
        self._add_file_row(
            card,
            row=2,
            label="결과 저장 경로",
            variable=self.output_path,
            command=self.select_output,
        )

        self.progress_bar = ttk.Progressbar(
            card,
            variable=self.progress_value,
            maximum=100,
            mode="determinate",
            style="Blue.Horizontal.TProgressbar",
        )
        self.progress_bar.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(18, 8))

        ttk.Label(card, textvariable=self.status_text, anchor="w", style="Status.TLabel").grid(
            row=4, column=0, columnspan=2, sticky="ew"
        )

        self.run_button = ttk.Button(
            card,
            text="실행",
            command=self.run,
            width=14,
            style="Primary.TButton",
        )
        self.run_button.grid(row=4, column=2, sticky="e", pady=(4, 0))

        container.grid_columnconfigure(0, weight=1)
        card.grid_columnconfigure(1, weight=1)

    def _add_file_row(self, parent, row, label, variable, command):
        ttk.Label(parent, text=label, anchor="w", width=18, style="Field.TLabel").grid(
            row=row, column=0, sticky="w", pady=8
        )
        ttk.Entry(parent, textvariable=variable, width=72).grid(
            row=row, column=1, sticky="ew", padx=(8, 8), pady=8
        )
        ttk.Button(parent, text="찾기", command=command, width=10, style="Browse.TButton").grid(
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

    def update_progress(self, percent, message):
        self.after(0, self._set_progress, percent, message)

    def _set_progress(self, percent, message):
        self.progress_value.set(percent)
        self.status_text.set(message)

    def set_running(self, is_running):
        state = tk.DISABLED if is_running else tk.NORMAL
        if self.run_button:
            self.run_button.config(state=state)

    def run(self):
        try:
            template_path, export_path, output_path = self.validate_inputs()
        except Exception as exc:
            messagebox.showerror("오류", str(exc))
            return

        self.progress_value.set(0)
        self.status_text.set("시작하는 중...")
        self.set_running(True)

        worker = threading.Thread(
            target=self._run_worker,
            args=(template_path, export_path, output_path),
            daemon=True,
        )
        worker.start()

    def _run_worker(self, template_path, export_path, output_path):
        try:
            copy_export_to_gl_input(
                template_path,
                export_path,
                output_path,
                progress_callback=self.update_progress,
            )
            self.after(0, self._show_success, output_path)
        except Exception as exc:
            self.after(0, self._show_error, str(exc))

    def _show_success(self, output_path):
        self.set_running(False)
        self.progress_value.set(100)
        self.status_text.set("완료")
        messagebox.showinfo("완료", f"결과 파일을 저장했습니다.\n\n{output_path}")

    def _show_error(self, message):
        self.set_running(False)
        self.status_text.set("오류 발생")
        messagebox.showerror("오류", message)


if __name__ == "__main__":
    app = GlInputCopyApp()
    app.mainloop()
