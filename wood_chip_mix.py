from openpyxl import load_workbook
import pandas as pd
from random import shuffle


EXCEL_PATH = "미이용산림_민호.xlsx"
SHEET_NAME = "Sheet1"
TARGET_WEIGHT = 1300

wb = load_workbook(EXCEL_PATH, data_only=True)
ws = wb.active

# 색이 없는 행만 수집
rows = []
for row in ws.iter_rows():  # 5행까지는 제외했으니 6행부터 시작
    cell_o = row[14]  # 'O'열 = 15번째 열 → 인덱스 14
    cell_t = row[19]  # 'T'열 = 20번째 열 → 인덱스 19

    # 배경색 확인
    fill_o = cell_o.fill.start_color.rgb
    fill_t = cell_t.fill.start_color.rgb

    # 색이 없는 셀만 선택 (흰색 또는 투명)
    if fill_o in ('00000000', 'FFFFFFFF', '00FFFFFF') and fill_t in ('00000000', 'FFFFFFFF', '00FFFFFF'):
        rows.append([cell_o.value, cell_t.value])


# 데이터 불러오기
df = pd.DataFrame(rows, columns=['발급번호', '목재칩'])


selected_df = df.loc[4:]

selected_df = selected_df.rename(columns={'Unnamed: 14': '발급번호', 'Unnamed: 19': '목재칩'})

selected_df['목재칩'] = pd.to_numeric(selected_df['목재칩'], errors='coerce')
filtered_df = selected_df.dropna(subset=['목재칩'])

items = list(zip(filtered_df['발급번호'], filtered_df['목재칩'].round(2)))

# 무게순 정렬 (가벼운 상품 먼저)
items.sort(key=lambda x: x[1])

print(items)

# 조합 함수: 중복 없이 가능한 많은 조합 생성
def pack_combos(items, target=1300):
    used = set()
    combos = []

    remaining = items.copy()
    while remaining:
        current_combo = []
        current_weight = 0

        for name, weight in remaining:
            if name in used:
                continue
            if current_weight + weight <= target * 1.5:  # 너무 큰 조합 제한
                current_combo.append((name, weight))
                current_weight += weight
                used.add(name)
            if current_weight >= target:
                break

        if current_weight >= target:
            combos.append((current_combo, current_weight))

        remaining = [item for item in items if item[0] not in used]
        if sum(w for name, w in remaining) < target:
            break

    return combos

def pack_best_fit_combos(items, target=1300):
    items = sorted(items, key=lambda x: x[1])  # 가벼운 순 정렬
    used = set()
    combos = []

    remaining = [item for item in items if item[0] not in used]

    while remaining:
        best_combo = None
        best_weight = float('inf')

        # 모든 가능한 조합 탐색
        for i in range(len(remaining)):
            combo = []
            total = 0
            for j in range(i, len(remaining)):
                name, weight = remaining[j]
                if name in used:
                    continue
                if total + weight > target * 1.5:  # 너무 무거운 조합 제외
                    break
                combo.append((name, weight))
                total += weight
                if total >= target:
                    break

            # 조합 후보가 타겟 이상일 경우, 더 좋은지 비교
            if total >= target and total < best_weight:
                best_combo = combo
                best_weight = total

        # 최적 조합이 있으면 추가
        if best_combo:
            combos.append((best_combo, best_weight))
            for name, _ in best_combo:
                used.add(name)
        else:
            break

        # 다음 반복을 위해 남은 상품 다시 필터링
        remaining = [item for item in items if item[0] not in used]

    return combos

# 조합 실행
grouped_combos = pack_best_fit_combos(items, TARGET_WEIGHT)

# 결과 출력
print("\n====== 전체 조합 결과 요약 ======\n")
all_used = set()

for i, (combo, weight) in enumerate(grouped_combos, 1):
    print(f"[조합 {i}] 총 무게: {weight}g / 상품 수: {len(combo)}개")
    for name, w in combo:
        print(f" - {name} ({w}g)")
        all_used.add(name)
    print()

# 조합된 상품과 조합되지 않은 상품 표시
print("✅ 조합에 포함된 모든 상품:")
print(", ".join(sorted(all_used)))

unused_items = [name for name, w in items if name not in all_used]
if unused_items:
    print("\n❌ 조합되지 않은 상품:")
    print(", ".join(sorted(unused_items)))
else:
    print("\n🎉 모든 상품이 조합에 사용되었습니다!")

#
# # 결과 엑셀에 저장
# result_df = pd.DataFrame(best_combo, columns=["상품명", "무게"])
# result_df.loc['합계'] = ['총합', best_weight]
#
# with pd.ExcelWriter(EXCEL_PATH, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
#     result_df.to_excel(writer, sheet_name='선택결과', index=False)
