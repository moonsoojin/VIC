import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re

# 함수: 파일 불러오기

def load_excel():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        excel_path.set(filepath)


def load_xy():
    filepath = filedialog.askopenfilename(filetypes=[("XY files", "*.xy")])
    if filepath:
        xy_path.set(filepath)


def save_xy():
    filepath = filedialog.asksaveasfilename(defaultextension=".xy", filetypes=[("XY files", "*.xy")])
    if filepath:
        save_path.set(filepath)


def update_tsdemand():
    try:
        # 엑셀 파일 로드 및 정리
        df = pd.read_excel(excel_path.get())
        df.columns = df.columns.astype(str)
        df.set_index('date', inplace=True)
        df = df.round().astype(int)

        # 노드명 필터링: A_로 시작하고 A_Saemangeum 제외
        valid_nodes = [col for col in df.columns if col.startswith("A_") and col != "A_Saemangeum"]

        # xy 파일 읽기
        with open(xy_path.get(), 'r', encoding='utf-8') as file:
            lines = file.readlines()

        new_lines = []
        i = 0
        current_node = None
        while i < len(lines):
            line = lines[i]
            match = re.match(r"\s*name\s*=\s*(A_[\w_]+)", line.strip())
            if match:
                current_node = match.group(1)

            # tsdemand 구간 수정
            if current_node in valid_nodes and line.strip().lower() == 'tsdemand':
                new_lines.append(line)  # 'tsdemand' 줄
                i += 1
                new_lines.append(lines[i])  # 'units' 줄
                i += 1
                # 날짜 + 값 교체
                for j in range(72):
                    date_str = lines[i].split("\t")[0].strip()
                    new_val = df[current_node].iloc[j]
                    new_lines.append(f"{date_str}\t{new_val}\n")
                    i += 1
                continue

            new_lines.append(line)
            i += 1

        # 저장
        with open(save_path.get(), 'w', encoding='utf-8') as file:
            file.writelines(new_lines)

        messagebox.showinfo("완료", "XY 파일이 성공적으로 저장되었습니다.")

    except Exception as e:
        messagebox.showerror("오류", str(e))


# GUI 구성
root = tk.Tk()
root.title("A_ 노드 tsdemand 일괄 수정기")
root.geometry("600x250")

excel_path = tk.StringVar()
xy_path = tk.StringVar()
save_path = tk.StringVar()

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(fill="both", expand=True)

# 엑셀 파일 선택
tk.Label(frame, text="Excel 입력 파일:").grid(row=0, column=0, sticky='e')
tk.Entry(frame, textvariable=excel_path, width=50).grid(row=0, column=1)
tk.Button(frame, text="불러오기", command=load_excel).grid(row=0, column=2)

# xy 파일 선택
tk.Label(frame, text="원본 XY 파일:").grid(row=1, column=0, sticky='e')
tk.Entry(frame, textvariable=xy_path, width=50).grid(row=1, column=1)
tk.Button(frame, text="불러오기", command=load_xy).grid(row=1, column=2)

# 저장 위치 설정
tk.Label(frame, text="저장될 XY 파일:").grid(row=2, column=0, sticky='e')
tk.Entry(frame, textvariable=save_path, width=50).grid(row=2, column=1)
tk.Button(frame, text="저장 위치", command=save_xy).grid(row=2, column=2)

# 실행 버튼
tk.Button(frame, text="tsdemand 시계열 수정", command=update_tsdemand, bg='green', fg='white').grid(row=4, column=1, pady=20)

root.mainloop()
