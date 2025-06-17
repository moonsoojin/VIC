import pandas as pd
import os
from glob import glob

# 1. 실제 파일 경로 입력
folder = "C:/Users/사용자명/Documents/04-Tank"  # 여기를 실제 경로로 수정하세요
files = glob(os.path.join(folder, "*.xlsx"))

# 2. 그룹 분류 기준
groups = {
    "농업_시군": [],
    "농업_표준": [],
    "생공_시군": [],
    "생공_표준": []
}

# 3. 파일 자동 분류
for f in files:
    filename = os.path.basename(f)
    for key in groups:
        if key in filename:
            groups[key].append(f)

# 4. 병합 함수 정의
def merge_group(file_list):
    dfs = []
    for f in file_list:
        if "Demand" in f:
            suffix = "_Demand"
        elif "Supply" in f:
            suffix = "_Supply"
        elif "Shortage" in f:
            suffix = "_Shortage"
        else:
            suffix = "_Unknown"

        df = pd.read_excel(f)
        df.reset_index(drop=True, inplace=True)
        key_col = df.columns[0]
        df[key_col] = df[key_col].astype(str)
        df.columns = [f"{col}{suffix}" if col != key_col else col for col in df.columns]
        dfs.append(df)

    base = dfs[0]
    for other in dfs[1:]:
        base = pd.merge(base, other, on=base.columns[0], how='outer')
    return base

# 5. 결과 저장
output_path = os.path.join(folder, "병합결과_4그룹.xlsx")
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    for group_name, file_list in groups.items():
        if len(file_list) == 3:
            merged_df = merge_group(file_list)
            merged_df.to_excel(writer, sheet_name=group_name[:31], index=False)
        else:
            print(f"[주의] {group_name} → 포함된 파일 수: {len(file_list)}개 (3개 필요)")

print(f"✅ 병합 완료! 결과 파일: {output_path}")
