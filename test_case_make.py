import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


def create_test_excel_file(file_name="test_data_100_items.xlsx", num_items=100):
    """
    테스트를 위한 가상 엑셀 파일을 생성합니다.
    A열(구분), O열(발급번호), T열(무게) 데이터를 포함합니다.
    """
    data = []
    categories = ['Category_A', 'Category_B', 'Category_C', 'Category_D', 'Category_E']

    for i in range(num_items):
        # A열 (구분) 데이터
        category = np.random.choice(categories)

        # O열 (발급번호) 데이터
        issue_num = f"WOOD_CHIP_{i + 1:03d}"

        # T열 (목재칩/무게) 데이터: 50g ~ 300g 사이의 무작위 값 (소수점 2자리)
        weight = round(np.random.uniform(50, 300), 2)

        data.append({
            'A_COL': category,
            'O_COL': issue_num,
            'T_COL': weight
        })

    df = pd.DataFrame(data)

    # openpyxl을 사용하여 엑셀 워크북 생성
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # 1행부터 5행까지는 비워두거나 더미 데이터 삽입 (코드의 스킵 조건 맞추기 위해)
    for r in range(1, 6):
        ws.cell(row=r, column=1, value=f"Header_{r}")
        ws.cell(row=r, column=15, value=f"Header_{r}")  # O열에도 더미
        ws.cell(row=r, column=20, value=f"Header_{r}")  # T열에도 더미

    # 데이터프레임의 내용을 엑셀에 쓰기
    # A열 (col 1), O열 (col 15), T열 (col 20)에 데이터 삽입
    start_row_for_data = 6  # 데이터가 시작될 행

    for r_idx, row_data in df.iterrows():
        ws.cell(row=start_row_for_data + r_idx, column=1, value=row_data['A_COL'])  # A열
        ws.cell(row=start_row_for_data + r_idx, column=15, value=row_data['O_COL'])  # O열
        ws.cell(row=start_row_for_data + r_idx, column=20, value=row_data['T_COL'])  # T열

    # 각 열의 너비 자동 조정 (선택 사항)
    ws.column_dimensions[get_column_letter(1)].width = 15  # A열
    ws.column_dimensions[get_column_letter(15)].width = 20  # O열
    ws.column_dimensions[get_column_letter(20)].width = 15  # T열

    wb.save(file_name)
    print(f"'{file_name}' 파일이 성공적으로 생성되었습니다. {num_items}개의 항목이 포함되어 있습니다.")


if __name__ == "__main__":
    create_test_excel_file()