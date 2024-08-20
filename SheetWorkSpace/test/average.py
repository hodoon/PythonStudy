import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# 맑은 고딕 폰트 파일 경로
font_path = 'C:\\Windows\\Fonts\\malgun.ttf'
font_prop = fm.FontProperties(fname=font_path)
plt.rcParams['font.family'] = font_prop.get_name()

# 1. 커스텀 평균 계산 함수 정의
def custom_average_columnwise(group):
    # 각 열에 대해 정렬 후, 최소값과 최대값을 제외한 나머지 값들의 평균을 계산 (소수점 2자리 반올림)
    return group.iloc[:, 2:].apply(lambda col: round(col.sort_values().iloc[1:-1].mean(), 2))

# 2. 파일 처리 함수 정의
def process_files(file_path, output_path):
    try:
        # 엑셀 파일 불러오기
        df = pd.read_excel(file_path)

        # '이름'과 '학번'을 기준으로 그룹화
        grouped = df.groupby(['이름', '학번'], group_keys=False)

        # 그룹화된 데이터에 대해 함수 적용 및 인덱스 재설정
        result = grouped.apply(custom_average_columnwise).reset_index()

        # 열 이름 변경 (명확성을 위해)
        result.columns = ['이름', '학번'] + [f'{col}' for col in df.columns[2:]]

        # 결과를 새로운 엑셀 파일로 저장
        result.to_excel(output_path, index=False)

        # 성공 메시지 표시
        messagebox.showinfo("성공", f"평균값이 저장된 파일 : {output_path}")
    except Exception as e:
        # 오류 메시지 표시
        messagebox.showerror("오류", f"파일을 저장하는 도중 오류가 발생했습니다 : {e}")

# 3. 파일 선택 함수 정의
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    return file_path

# 4. GUI 구성
def main():
    # GUI 창 설정
    root = tk.Tk()
    root.withdraw # 기본 GUT 창 숨기기

    # 'my_view.xlsx' 파일 선택
    messagebox.showinfo("파일 선택", "평균을 낼 취업캠프 설문조사 결과 파일을 선택하세요.")
    file_path = select_file()

    if not file_path:
        messagebox.showinfo("경고", "엑셀 파일이 선택되지 않았습니다.")
        return

    # 출력 엑셀 파일 경로 선택
    messagebox.showinfo("파일 저장", "저장할 엑셀 파일 경로를 선택하세요.")
    output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

    if not output_path:
        messagebox.showwarning("경고", "저장할 경로가 선택되지 않았습니다.")
        return

    # 선택된 파일 경로 및 디렉토리를 사용하여 파일 처리
    process_files(file_path, output_path)

if __name__ == "__main__":
    main()




