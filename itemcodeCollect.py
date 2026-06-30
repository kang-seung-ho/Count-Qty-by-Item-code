import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


def normalize_product_code(value):
    """
    상품코드 정규화
    - 숫자형 123.0 → 123
    - 텍스트형 '00123' → 00123 유지
    - 공백 제거
    """
    if pd.isna(value):
        return ""

    text = str(value).strip()

    # 엑셀 숫자값이 123.0처럼 읽힌 경우 123으로 변환
    if text.endswith(".0"):
        text = text[:-2]

    return text


def normalize_quantity(value):
    """
    수량 정규화
    - 숫자형 대응
    - 텍스트 숫자 대응
    - '1,000' 대응
    - 빈 값/오류 값은 0 처리
    """
    if pd.isna(value):
        return 0

    text = str(value).strip()
    text = text.replace(",", "")

    qty = pd.to_numeric(text, errors="coerce")

    if pd.isna(qty):
        return 0

    return qty


def select_file():
    file_path = filedialog.askopenfilename(
        title="엑셀 파일 선택",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not file_path:
        return

    try:
        # dtype=object: 엑셀의 텍스트/숫자 서식을 최대한 원본 형태로 읽기
        df = pd.read_excel(file_path, dtype=object)

        # A열, B열만 사용
        df = df.iloc[:, [0, 1]]
        df.columns = ["상품코드", "수량"]

        # 상품코드/수량 정리
        df["상품코드"] = df["상품코드"].apply(normalize_product_code)
        df["수량"] = df["수량"].apply(normalize_quantity)

        # 상품코드 없는 행 제거
        df = df[df["상품코드"] != ""]

        # 상품코드별 수량 합계
        result = df.groupby("상품코드", as_index=False)["수량"].sum()

        # 수량이 정수로 떨어지면 정수 처리
        result["수량"] = result["수량"].apply(
            lambda x: int(x) if float(x).is_integer() else x
        )

        save_path = filedialog.asksaveasfilename(
            title="결과 엑셀 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not save_path:
            return

        result.to_excel(save_path, index=False)

        messagebox.showinfo("완료", "상품별 수량 합계 엑셀이 저장되었습니다.")

    except Exception as e:
        messagebox.showerror("오류", f"처리 중 오류가 발생했습니다.\n\n{e}")


root = tk.Tk()
root.title("상품별 수량 합계 프로그램")
root.geometry("480x200")

label = tk.Label(
    root,
    text="엑셀 파일을 선택하면\n상품코드별 수량 합계를 계산합니다.\n A열헤더에 '상품코드', B열헤더에 '수량'을 기재해야 합니다.",
    font=("맑은 고딕", 11)
)
label.pack(pady=20)

button = tk.Button(
    root,
    text="엑셀 파일 선택",
    command=select_file,
    width=22,
    height=2
)
button.pack()

root.mainloop()