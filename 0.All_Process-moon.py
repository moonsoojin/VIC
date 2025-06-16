import os
import sys
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import pyodbc
import time
import csv

os.chdir(os.path.dirname(os.path.abspath(__file__)))

def select_file(filetypes, title="Select a file"):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    root.destroy()
    return file_path

def extract_mdb():
    print("\n [1] MDB 파일 추출 시작")
    mdb_path = select_file([("Access Database", "*.mdb;*.accdb")], "MDB 파일을 선택하세요")
    if not mdb_path:
        print("❌ MDB 파일을 선택하지 않았습니다.")
        return None
    
    odbc_driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
    conn_str = f'DRIVER={odbc_driver};DBQ={mdb_path}'
    filename = os.path.basename(mdb_path).replace('OUTPUT.mdb', '')
    output_dir = os.path.join(os.path.dirname(mdb_path), "1.mdb2txt")
    os.makedirs(output_dir, exist_ok=True)

    table_names = ["DEMOutput", "NodesInfo", "TimeSteps"]
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        for table in table_names:
            try:
                cursor.execute(f"SELECT * FROM {table}")
                rows = cursor.fetchall()
                columns = [col[0] for col in cursor.description]
                csv_path = os.path.join(output_dir, f"{filename}_{table}.txt")
                with open(csv_path, 'w', encoding='utf-8') as f:
                    writer = csv.writer(f, lineterminator='\n')
                    writer.writerow(columns)
                    writer.writerows(rows)
                print(f" {csv_path} 저장 완료.")
            except Exception as e:
                print(f"❌ {table} 테이블 처리 실패: {e}")
        conn.close()
    except Exception as e:
        print(f"❌ MDB 처리 실패: {e}")
        return None
    time.sleep(1)
    # DEMOutput.txt 파일 경로 자동 생성 (확장자 추가)
    dem_output_path = os.path.join(output_dir, f"{filename}_DEMOutput.txt")
    return output_dir, dem_output_path

def merge_txt(output_dir, dem_output_file):
    print("\n [2] TXT 병합 시작")
    if not dem_output_file:
        print("❌ DEMOutput 파일을 선택하지 않았습니다.")
        return None
    
    if not os.path.exists(dem_output_file):
        print(f"❌ DEMOutput 파일이 존재하지 않습니다: {dem_output_file}")
        return None
    
    parent_dir = os.path.dirname(output_dir)
    results_dir = os.path.join(parent_dir, "2.extract_Node_5days")
    os.makedirs(results_dir, exist_ok=True)
    
    prefix = os.path.basename(dem_output_file).split('DEMOutput.txt')[0]
    
    df1 = pd.read_csv(dem_output_file)
    df2 = pd.read_csv(os.path.join(output_dir, f"{prefix}NodesInfo.txt"))
    df3 = pd.read_csv(os.path.join(output_dir, f"{prefix}TimeSteps.txt"))
    
    df1.rename(columns={'NNo':'NNumber'}, inplace=True)
    df4 = pd.merge(df1, df2)
    df4.drop(labels=['Gw_In', 'Hydro_State'], axis=1, inplace=True)
    df5 = pd.merge(df3, df4)
    df5.drop(labels=['MidDate', 'Duration', 'EndDate'], axis=1, inplace=True)
    df6 = df5[df5['NType'].str.startswith('Demand')]
    df6 = df6[~df6['NName'].str.startswith(('F', '@'))]
    df6.drop(labels=['NType'], axis=1, inplace=True)
    df7 = df6.groupby(['TSDate', 'NName'])[['Shortage', 'Surf_In', 'Demand']].sum().reset_index()
    
    df7 = pd.concat([df7[df7['NName'].str.startswith(prefix)] for prefix in ['A_', 'D_', 'I_', 'DI_', 'D_Mi_']])
    
    df7 = df7.groupby(['TSDate', 'NName'])[['Shortage', 'Surf_In', 'Demand']].sum().reset_index()
    Table1 = df7.pivot_table(values='Shortage', index='TSDate', columns='NName')
    Table2 = df7.pivot_table(values='Demand', index='TSDate', columns='NName')
    Table3 = df7.pivot_table(values='Surf_In', index='TSDate', columns='NName')
    output_file1 = os.path.join(results_dir, os.path.basename(dem_output_file).replace('DEMOutput.txt', 'Shortage.csv'))
    output_file2 = os.path.join(results_dir, os.path.basename(dem_output_file).replace('DEMOutput.txt', 'Demand.csv'))
    output_file3 = os.path.join(results_dir, os.path.basename(dem_output_file).replace('DEMOutput.txt', 'Supply.csv'))
    Table1.to_csv(output_file1, encoding='utf-8-sig')
    Table2.to_csv(output_file2, encoding='utf-8-sig')
    Table3.to_csv(output_file3, encoding='utf-8-sig')
    print(f"병합 결과 저장: {output_file1}")
    print(f"생성된 DEMOutput 파일 경로: {dem_output_file}")
    return {'Shortage': output_file1, 'Demand': output_file2, 'Supply':output_file3}

def area_ratio_conversion(result_file, ratio_csv, label):
    print(f"\n [3] {label} 급수비율 계산 시작")
    df_ratio = pd.read_csv(ratio_csv)
    df_ratio = df_ratio.loc[:, ~df_ratio.columns.str.contains('^Unnamed')]  # <-- 이 줄 추가
    parent_dir = os.path.dirname(os.path.dirname(result_file))
    result_dir = os.path.join(parent_dir, "3.ratio_results")
    os.makedirs(result_dir, exist_ok=True)
    
    df_ratio.rename(columns={df_ratio.columns[0]: '노드'}, inplace=True)
    df_ratio.set_index('노드', inplace=True)
    df_ratio = df_ratio.astype(float)
    
    df_results = pd.read_csv(result_file)
    time_col = df_results['TSDate'] if 'TSDate' in df_results.columns else None
    df_numeric = df_results.drop(columns=['TSDate']) if time_col is not None else df_results
    
    common_nodes = df_ratio.index.intersection(df_numeric.columns)
    df_product = df_numeric[common_nodes].dot(df_ratio.loc[common_nodes])
    df_final = pd.DataFrame(df_product, columns=df_ratio.columns)
    if time_col is not None:
        df_final.insert(0, 'TSDate', time_col)
    
    basename = os.path.splitext(os.path.basename(result_file))[0]
    out_path = os.path.join(result_dir, f"{basename}_{label}.csv")
    df_final.to_csv(out_path, index=False, encoding='utf-8-sig')
    print(f"결과 저장: {out_path}")
    return out_path

def convert_to_yearly(input_csv):
    print(f"\n [4] 연도별 변환 시작")
    parent_dir = os.path.dirname(os.path.dirname(input_csv))
    result_dir = os.path.join(parent_dir, "4.ratio_results_wateryear")
    os.makedirs(result_dir, exist_ok=True)

    df = pd.read_csv(input_csv)
    df['TSDate'] = pd.to_datetime(df['TSDate'])
    df['WaterYear'] = df['TSDate'].apply(lambda x: x.year + 1 if x.month >= 10 else x.year).astype(int)
    # grouped = df.groupby('TSDate').sum().reset_index().set_index('TSDate').T
    df_numeric = df.drop(columns=['TSDate'])
    grouped = df_numeric.groupby('WaterYear').sum().T
    grouped.columns.name = None
    
    output_file = os.path.join(result_dir, os.path.basename(input_csv).replace(".csv", "_year.csv"))
    grouped.to_csv(output_file, encoding='utf-8-sig')
    print(f"연도별 결과 저장: {output_file}")
    return output_file

def merge_yearly(intake_csv, irrigation_csv, base_filename):
    print("\n [5] 생공 + 농업 연도별 결과 병합 시작")

    df1 = pd.read_csv(intake_csv, index_col=0)
    df2 = pd.read_csv(irrigation_csv, index_col=0)

    df1.index.name = '표준유역번호'
    df2.index.name = '표준유역번호'

    df1_sum = df1.sum(axis=1)
    df2_sum = df2.sum(axis=1)

    # 데이터프레임 결합
    df_merged = pd.DataFrame({
        '생공': df1_sum,
        '농업': df2_sum
    })

    df_merged['합계'] = df_merged['생공'] + df_merged['농업']
    df_merged.reset_index(inplace=True)

    # 저장
    parent_dir = os.path.dirname(os.path.dirname(intake_csv))
    output_dir = os.path.join(parent_dir, "4.result_merge_yearly")
    os.makedirs(output_dir, exist_ok=True)

    merged_file = os.path.join(output_dir, f"{base_filename}_result_생공+농업.csv")
    df_merged.to_csv(merged_file, index=False, encoding='utf-8-sig')

    print(f"생공 + 농업 합계 결과 저장 완료: {merged_file}")

def main():
    output_dir, dem_output_file = extract_mdb()
    if not output_dir:
        return

    result_csv = merge_txt(output_dir, dem_output_file)
    
    ratio_configs = [
        ('급수비율_생공_표준유역.csv', '생공_표준'),
        ('급수비율_농업_표준유역.csv', '농업_표준'),
        ('급수비율_생공_시군별.csv', '생공_시군'),
        ('급수비율_농업_시군별.csv', '농업_시군'),
        ]
    
    
    yearly_outputs = {}  # 연도별 결과 저장
    for label, path in result_csv.items():  # Shortage, Demand, Supply
      yearly_outputs[label] = {}
      for ratio_file, ratio_label in ratio_configs:
        print(f"⏳ [{label}] + [{ratio_label}] 처리 중...")
        converted = area_ratio_conversion(path, ratio_file, ratio_label)
        yearly = convert_to_yearly(converted)
        yearly_outputs[label][ratio_label] = yearly

if __name__ == "__main__":
    main()
