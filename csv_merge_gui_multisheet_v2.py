import os
import pandas as pd
from glob import glob
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

selected_folder = None  # 전역 변수 설정

def merge_files(file_list):
    dfs = []
    for f in file_list:
        try:
            df = pd.read_csv(f, encoding='utf-8')
        except:
            df = pd.read_csv(f, encoding='cp949')
        df.reset_index(drop=True, inplace=True)
        key_col = df.columns[0]
        df[key_col] = df[key_col].astype(str)
        suffix = "_Demand" if "Demand" in f else "_Supply" if "Supply" in f else "_Shortage" if "Shortage" in f else "_X"
        df.columns = [key_col if i == 0 else f"{col}{suffix}" for i, col in enumerate(df.columns)]
        dfs.append(df)

    if not dfs:
        return pd.DataFrame()

    merged = dfs[0]
    for df in dfs[1:]:
        merged = pd.merge(merged, df, on=merged.columns[0], how='outer')
    return merged

def merge_csv_files(folder, status_label, progress_var, progress_bar):
    if not folder or not os.path.isdir(folder):
        messagebox.showerror("오류", "올바른 폴더가 선택되지 않았습니다.")
        return

    files = glob(os.path.join(folder, "*.csv"))
    if not files:
        messagebox.showwarning("경고", f"'{folder}' 폴더에 .csv 파일이 없습니다.")
        return

    grouped_files = {
        "농업_시군": [],
        "농업_표준": [],
        "생공_시군": [],
        "생공_표준": []
    }

    for f in files:
        for group in grouped_files:
            if group in os.path.basename(f):
                grouped_files[group].append(f)

    total_steps = sum(len(file_list) for file_list in grouped_files.values())
    progress_var.set(0)
    progress_bar.config(maximum=total_steps)
    step = 0

    output_path = os.path.join(folder, "병합결과_여러시트.xlsx")
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for group, file_list in grouped_files.items():
            if len(file_list) >= 2:
                status_label.config(text=f"[{group}] 병합 중...", fg="blue")
                status_label.update_idletasks()
                merged_df = merge_files(file_list)
                if not merged_df.empty:
                    safe_sheet = group[:31]
                    merged_df.to_excel(writer, sheet_name=safe_sheet, index=False)
                step += len(file_list)
                progress_var.set(step)
                progress_bar.update_idletasks()
            else:
                status_label.config(text=f"[{group}] 파일 수 부족 ({len(file_list)})", fg="red")
                status_label.update_idletasks()

    progress_var.set(total_steps)
    status_label.config(text=f"✅ 병합 완료! 결과 파일:\n{output_path}", fg="green")
    try:
        os.startfile(output_path)
    except Exception:
        messagebox.showinfo("완료", f"엑셀 파일 경로:\n{output_path}")

def select_folder(status_label, folder_label):
    global selected_folder
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        selected_folder = folder_selected
        folder_label.config(text=f"선택된 폴더: {selected_folder}", fg="blue")
        status_label.config(text="이제 '병합' 버튼을 누르세요.", fg="black")

def start_merge(status_label, progress_var, progress_bar):
    if not selected_folder:
        messagebox.showwarning("알림", "먼저 CSV 폴더를 선택하세요.")
        return
    status_label.config(text="병합 중입니다. 잠시만 기다리세요...", fg="blue")
    threading.Thread(target=merge_csv_files, args=(selected_folder, status_label, progress_var, progress_bar), daemon=True).start()

def main():
    global selected_folder
    selected_folder = None
    root = tk.Tk()
    root.title("CSV 병합 도우미 (여러 시트)")
    root.geometry("540x340")
    root.resizable(False, False)

    label = tk.Label(root, text="CSV 파일 폴더 선택 후 '병합'을 누르세요.\n(그룹별 병합, 시트로 구분)", font=("맑은 고딕", 12))
    label.pack(pady=12)

    folder_label = tk.Label(root, text="선택된 폴더가 없습니다.", font=("맑은 고딕", 10))
    folder_label.pack(pady=5)

    status_label = tk.Label(root, text="", font=("맑은 고딕", 10))
    status_label.pack(pady=5)

    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=420)
    progress_bar.pack(pady=10)

    btn_select = tk.Button(root, text="CSV 폴더 선택", command=lambda: select_folder(status_label, folder_label), width=18, height=2)
    btn_select.pack(pady=5)

    btn_merge = tk.Button(root, text="병합", command=lambda: start_merge(status_label, progress_var, progress_bar), width=18, height=2)
    btn_merge.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()
