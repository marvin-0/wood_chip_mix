import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from tkinter import font

def pack_best_fit_combos(items, target=1300):
    """
    주어진 목표 무게에 가장 가깝게 아이템을 조합합니다.
    items는 [(A열값, 발급번호, 목재칩), ...] 형태를 예상합니다.
    """
    items = sorted(items, key=lambda x: x[2])
    used = set()
    combos = []

    remaining = [item for item in items if item[1] not in used]

    while remaining:
        best_combo = None
        best_weight = float('inf')

        for i in range(len(remaining)):
            combo = []
            total = 0
            for j in range(i, len(remaining)):
                if j >= len(remaining):
                    break
                a_value, issue_num, weight = remaining[j]
                if issue_num in used:
                    continue
                if total + weight > target * 1.5:
                    break
                combo.append((a_value, issue_num, weight))
                total += weight
                if total >= target:
                    break

            if total >= target and total < best_weight:
                best_combo = combo
                best_weight = total

        if best_combo:
            combos.append((best_combo, best_weight))
            for a_val, issue_num, _ in best_combo:
                used.add(issue_num)
        else:
            break

        remaining = [item for item in items if item[1] not in used]
    return combos

class TimberChipCombinerApp:
    def __init__(self, master):
        self.master = master
        master.title("목재칩 조합기")

        # 창 크기 조절에 따라 위젯들이 확장되도록 설정은 유지
        master.grid_rowconfigure(4, weight=1) # 결과 텍스트 영역이 세로로 확장 (폰트 조절 UI 추가로 행 번호 변경)
        master.grid_columnconfigure(1, weight=1) # 파일 경로 Entry와 버튼이 가로로 확장

        self.excel_path = tk.StringVar()
        self.target_weight = tk.DoubleVar(value=1300.0)

        # 폰트 크기 조절을 위한 StringVar 및 기본값
        self.font_size_var = tk.StringVar(value="10") # 기본 폰트 크기 10

        # 파일 경로 선택
        tk.Label(master, text="엑셀 파일 경로:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        tk.Entry(master, textvariable=self.excel_path, width=50).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        tk.Button(master, text="찾아보기", command=self.browse_excel_file).grid(row=0, column=2, padx=5, pady=5, sticky="e")

        # 목표 무게 설정
        tk.Label(master, text="목표 무게 (g):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        tk.Entry(master, textvariable=self.target_weight, width=10).grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # 폰트 크기 설정 UI 추가 (새로운 행)
        tk.Label(master, text="글씨 크기:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        tk.Entry(master, textvariable=self.font_size_var, width=5).grid(row=2, column=1, padx=5, pady=5, sticky="w")
        tk.Button(master, text="적용", command=self.apply_font_size).grid(row=2, column=2, padx=5, pady=5, sticky="w")


        # 조합 버튼 (행 번호 변경)
        tk.Button(master, text="조합 실행", command=self.perform_combination).grid(row=3, column=0, columnspan=3, pady=10)

        # 결과 출력 영역 (행 번호 변경)
        self.result_text_font = font.Font(family="TkDefaultFont", size=int(self.font_size_var.get()))
        self.result_text = scrolledtext.ScrolledText(master, wrap=tk.WORD, font=self.result_text_font)
        self.result_text.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")

        # 결과 저장 버튼 (행 번호 변경)
        self.save_button = tk.Button(master, text="결과 저장", command=self.save_results)
        self.save_button.grid(row=5, column=0, columnspan=3, pady=10)
        self.save_button.config(state=tk.DISABLED)

        self.grouped_combos = []
        self.original_items_for_unused = []

    def apply_font_size(self):
        """사용자가 입력한 값으로 폰트 크기를 적용합니다."""
        try:
            new_size = int(self.font_size_var.get())
            if new_size < 1: # 폰트 크기는 최소 1 이상
                messagebox.showwarning("입력 오류", "글씨 크기는 1 이상의 정수여야 합니다.")
                self.font_size_var.set("10") # 기본값으로 되돌림
                new_size = 10
            self.result_text_font.config(size=new_size)
        except ValueError:
            messagebox.showwarning("입력 오류", "올바른 숫자(정수)를 입력해주세요.")
            self.font_size_var.set("10") # 잘못된 입력 시 기본값으로 되돌림

    def browse_excel_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_path.set(file_path)

    def load_and_filter_data(self, excel_path):
        try:
            wb = load_workbook(excel_path, data_only=True)
            ws = wb.active

            rows = []
            for row_idx, row in enumerate(ws.iter_rows()):
                if row_idx < 5:
                    continue

                cell_a = row[0]
                cell_o = row[14]
                cell_t = row[19]

                fill_o = cell_o.fill.start_color.rgb if cell_o.fill.start_color else None
                fill_t = cell_t.fill.start_color.rgb if cell_t.fill.start_color else None

                if (fill_o in ('00000000', 'FFFFFFFF', None)) and \
                   (fill_t in ('00000000', 'FFFFFFFF', None)):
                    rows.append([cell_a.value, cell_o.value, cell_t.value])

            df = pd.DataFrame(rows, columns=['구분', '발급번호', '목재칩'])

            df['목재칩'] = pd.to_numeric(df['목재칩'], errors='coerce')
            filtered_df = df.dropna(subset=['목재칩', '발급번호'])

            items = []
            for _, row in filtered_df.iterrows():
                items.append((row['구분'], row['발급번호'], round(row['목재칩'], 2)))

            return items

        except Exception as e:
            messagebox.showerror("파일 읽기 오류", f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
            return None

    def perform_combination(self):
        excel_path = self.excel_path.get()
        if not excel_path:
            messagebox.showwarning("입력 오류", "엑셀 파일 경로를 입력해주세요.")
            return

        target_weight = self.target_weight.get()
        if target_weight <= 0:
            messagebox.showwarning("입력 오류", "목표 무게는 0보다 커야 합니다.")
            return

        items = self.load_and_filter_data(excel_path)
        if not items:
            return

        self.original_items_for_unused = items

        self.result_text.delete(1.0, tk.END)
        self.grouped_combos = pack_best_fit_combos(items, target_weight)

        output = []
        all_used_issue_nums = set()

        output.append("\n====== 전체 조합 결과 요약 ======\n")
        if not self.grouped_combos:
            output.append("💡 목표 무게에 맞는 조합을 찾을 수 없습니다.\n")
        else:
            for i, (combo, weight) in enumerate(self.grouped_combos, 1):
                output.append(f"[조합 {i}] 총 무게: {weight:.2f}g / 상품 수: {len(combo)}개\n")
                for a_val, issue_num, w in combo:
                    output.append(f" - (구분: {a_val if a_val is not None else 'N/A'}) {issue_num} ({w:.2f}g)\n")
                    all_used_issue_nums.add(issue_num)
                output.append("\n")

            output.append("✅ 조합에 포함된 모든 상품 발급번호:\n")
            output.append(", ".join(sorted(all_used_issue_nums)) + "\n")

            unused_items_detail = [(a, issue, w) for a, issue, w in self.original_items_for_unused if issue not in all_used_issue_nums]
            if unused_items_detail:
                output.append("\n❌ 조합되지 않은 상품:\n")
                for a_val, issue_num, w in unused_items_detail:
                    output.append(f" - (구분: {a_val if a_val is not None else 'N/A'}) {issue_num} ({w:.2f}g)\n")
            else:
                output.append("\n🎉 모든 상품이 조합에 사용되었습니다!\n")

        self.result_text.insert(tk.END, "".join(output))
        self.save_button.config(state=tk.NORMAL)

    def save_results(self):
        if not self.grouped_combos:
            messagebox.showwarning("저장 오류", "저장할 조합 결과가 없습니다. 먼저 '조합 실행'을 해주세요.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="조합 결과 저장"
        )
        if not save_path:
            return

        try:
            wb_out = Workbook()
            ws_out = wb_out.active
            ws_out.title = "조합 결과"

            ws_out.append(["조합 정보", "구분", "발급번호", "무게 (g)"])
            for col in range(1, 5):
                ws_out.cell(row=1, column=col).fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

            row_idx = 2

            for i, (combo, weight) in enumerate(self.grouped_combos, 1):
                ws_out.cell(row=row_idx, column=1, value=f"[조합 {i}]")
                ws_out.cell(row=row_idx, column=2, value=f"총 무게: {weight:.2f}g")
                ws_out.cell(row=row_idx, column=3, value=f"상품 수: {len(combo)}개")
                row_idx += 1
                for a_val, issue_num, w in combo:
                    ws_out.cell(row=row_idx, column=1, value="")
                    ws_out.cell(row=row_idx, column=2, value=a_val)
                    ws_out.cell(row=row_idx, column=3, value=issue_num)
                    ws_out.cell(row=row_idx, column=4, value=f"{w:.2f}")
                    row_idx += 1
                row_idx += 1

            all_used_issue_nums = set()
            for combo, _ in self.grouped_combos:
                for a_val, issue_num, w in combo:
                    all_used_issue_nums.add(issue_num)

            used_categories = sorted(list(set([item[0] for item in self.original_items_for_unused if item[1] in all_used_issue_nums])))

            row_idx += 1
            ws_out.cell(row=row_idx, column=1, value="✅ 조합에 사용된 상품 (구분별 오름차순):")
            row_idx += 1
            for category in used_categories:
                ws_out.cell(row=row_idx, column=1, value=category)
                row_idx += 1
            row_idx += 1

            unused_categories = sorted(list(set([item[0] for item in self.original_items_for_unused if item[1] not in all_used_issue_nums])))

            ws_out.cell(row=row_idx, column=1, value="❌ 조합되지 않은 상품 (구분별 오름차순):")
            row_idx += 1
            for category in unused_categories:
                ws_out.cell(row=row_idx, column=1, value=category)
                row_idx += 1

            for col in range(1, ws_out.max_column + 1):
                ws_out.column_dimensions[get_column_letter(col)].width = 15
                max_length = 0
                for cell in ws_out[get_column_letter(col)]:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.2
                if adjusted_width > ws_out.column_dimensions[get_column_letter(col)].width:
                     ws_out.column_dimensions[get_column_letter(col)].width = adjusted_width

            wb_out.save(save_path)
            messagebox.showinfo("저장 완료", f"조합 결과가 '{save_path}'에 성공적으로 저장되었습니다.")
        except Exception as e:
            messagebox.showerror("저장 오류", f"파일 저장 중 오류가 발생했습니다: {e}")

root = tk.Tk()
app = TimberChipCombinerApp(root)
root.mainloop()