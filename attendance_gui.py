#!/usr/bin/env python3
"""
attendance_gui.py

GUI 기반 근태 자동정리기
- 연/월 입력
- 공장 파일 선택 (선택)
- 사무실 파일 선택 (선택)
- 보고 템플릿 파일 선택 (사용자 제공; 기본 템플릿 선택 가능)
- 출력 폴더 선택
- 실행 시: 기존 출력파일(근태결과_YYYYMM.xlsx)이 있으면 해당 시트만 갱신(입력한 쪽만), 없으면 생성
- 병합 규칙: 같은 성명+날짜 -> 출근=min(출근들), 퇴근=max(퇴근들)
"""

import sys, os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import shutil

import traceback

APP_VERSION = "v1.0.0"
DEVELOPER = "왕형순"
BUILD_DATE = "2025-10-23"

# ----------------- 데이터 처리 -----------------

def process_factory(file):
    if not file:
        return pd.DataFrame(columns=['성명','사번','부서','날짜','출근시간','퇴근시간'])
    try:
        # df = pd.read_excel(file)
        # 1) 일반적으로 읽기 (헤더가 복수라인이면 아래에서 재시도)
        # df = pd.read_excel(factory_path, engine='openpyxl', header=3, dtype=str)
        for i in range(5):
            df_tmp = pd.read_excel(file, engine='openpyxl', header=i)
            cols = [str(c).strip() for c in df_tmp.columns]
            if "출입날짜" in cols and "출입시간" in cols:
                df = df_tmp
                break
        else:
            raise ValueError("출입날짜 / 출입시간 컬럼을 찾을 수 없습니다.")
    except Exception as e:
        messagebox.showerror("파일 읽기 실패", f"공장 파일을 읽는 중 에러:\n{e}")
        return pd.DataFrame(columns=['성명','사번','부서','날짜','출근시간','퇴근시간'])

    # 모든 컬럼명에서 공백 제거
    df.columns = df.columns.str.strip()

    # 진단: 실제 컬럼 확인 (콘솔 출력 + 메시지박스 로그는 필요시 활용)
    print("[DEBUG] process_factory - columns:", df.columns.tolist())

    df = df[df['기능키'].isin(['출근','출입','퇴근'])].copy()
    df['날짜'] = pd.to_datetime(df['출입날짜'], format="%Y-%m-%d", errors='coerce').dt.date
    df['시간'] = pd.to_datetime(df['출입시간'], format="%H:%M:%S", errors='coerce').dt.time
    df['성명'] = df['이  름']
    df['사번'] = df['사  번']
    grouped = (
        df.groupby(['성명','사번','날짜'])
        .agg(출근시간=('시간','min'),
            퇴근시간=('시간','max'))
        .reset_index()
    )
    grouped['구분'] = '공장'
    return grouped

def process_office(file):
    if not file:
        return pd.DataFrame(columns=['성명','사번','부서','날짜','출근시간','퇴근시간'])
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
    except Exception as e:
        messagebox.showerror("파일 읽기 실패", f"사무실 파일을 읽는 중 에러:\n{e}")
        return pd.DataFrame(columns=['성명','사번','부서','날짜','출근시간','퇴근시간'])

    data = []
    for r in range(5, ws.max_row+1, 3):  # 실제 오프셋은 샘플에서 조정 필요
        name = ws.cell(r, 11).value
        emp_num = ws.cell(r, 3).value
        dept = ws.cell(r, 21).value
        if not name: continue
        for c in range(5, 36):
            time_stamp = None
            time_stamp_split = None
            start = None
            end = None
            date = c-4

            # start = ws.cell(r, c).value
            # end = ws.cell(r+1, c).value
            time_stamp = ws.cell(r+1, date).value
            if time_stamp is not None:
                time_stamp_split = time_stamp.splitlines()
                start = time_stamp_split[0]
                end = time_stamp_split[-1]

            if start or end:
                data.append([name, emp_num, dept, date, start, end])
    df = pd.DataFrame(data, columns=['성명','사번','부서','일','출근시간','퇴근시간'])

    # 모든 컬럼명에서 공백 제거
    df.columns = df.columns.str.strip()

    for col in ['출근시간','퇴근시간']:
        df[col] = pd.to_datetime(df[col], format="%H:%M", errors='coerce').dt.time
    df['날짜'] = pd.to_datetime(f"2025-10-" + df['일'].astype(str)).dt.date
    df = df.drop(columns=['부서','일'])
    df['구분'] = '사무실'
    df = df.reindex(columns=['성명','사번','날짜','출근시간','퇴근시간','구분'])
    return df

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def merge_data(factory_df=None, office_df=None):
    dfs = [df for df in [factory_df, office_df] if df is not None]
    if not dfs:
        return pd.DataFrame()
    df = pd.concat(dfs, ignore_index=True)
    df = df.sort_values(by=['성명','사번','날짜','출근시간'])
    df = df.groupby(['성명','사번','날짜'], as_index=False).agg({
        '출근시간':'min',
        '퇴근시간':'max',
        '구분':'first'
    })
    df['근무시간'] = (
        pd.to_datetime(df['퇴근시간'].astype(str)) -
        pd.to_datetime(df['출근시간'].astype(str))
    ).dt.total_seconds() / 3600
    return df

def update_excel(factory_df, office_df, merged_df, year_str, month_str):
    out_path = Path(f"output/근태기록_{year_str}_{month_str}.xlsx")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    template = Path(resource_path("template.xlsx"))

    if not out_path.exists():
        shutil.copy(template, out_path)

    wb = openpyxl.load_workbook(out_path)
    # 갱신 대상 시트만 재작성
    if factory_df is not None:
        if "공장데이터" in wb.sheetnames:
            ws = wb["공장데이터"]
            wb.remove(ws)
        ws = wb.create_sheet("공장데이터")
        for r in dataframe_to_rows(factory_df, index=False, header=True):
            ws.append(r)

    if office_df is not None:
        if "사무실데이터" in wb.sheetnames:
            ws = wb["사무실데이터"]
            wb.remove(ws)
        ws = wb.create_sheet("사무실데이터")
        for r in dataframe_to_rows(office_df, index=False, header=True):
            ws.append(r)

    if not merged_df.empty:
        if "정리테이블" in wb.sheetnames:
            ws = wb["정리테이블"]
            wb.remove(ws)
        ws = wb.create_sheet("정리테이블")
        for r in dataframe_to_rows(merged_df, index=False, header=True):
            ws.append(r)

        # 보고용 시트 갱신
        # if "보고용" in wb.sheetnames:
        #     ws = wb["보고용"]
        #     wb.remove(ws)
        # ws = wb.create_sheet("보고용")
        # ws["A1"] = f"{year_str}_{month_str} 근태 보고서"
        # for r_idx, row in enumerate(merged_df.itertuples(), start=3):
        #     ws.cell(r_idx, 1).value = row.성명
        #     ws.cell(r_idx, 2).value = row.날짜.strftime("%Y-%m-%d")
        #     ws.cell(r_idx, 3).value = str(row.출근시간)
        #     ws.cell(r_idx, 4).value = str(row.퇴근시간)
        #     ws.cell(r_idx, 5).value = round(row.근무시간, 2)

        print("year_str type" + type(year_str))

        report_sheet_name = str(year_str % 100) + '.' + str(month_str)
        if report_sheet_name in wb.sheetnames:
            ws = wb[report_sheet_name]
            wb.remove(ws)
        ws = wb.copy_worksheet("tmp")
        ws.title = report_sheet_name
        ws['A1'] = '㈜대영인텍 출퇴근기록부 - ' + str(year_str) + '년 ' + str(month_str) + '월'

        wb.move_sheet('tmp', 4)
        wb.move_sheet(report_sheet_name, -3)
        wb.move_sheet('정리테이블', 2)

    wb.save(out_path)
    return out_path

# ----------------- GUI 구성 -----------------

def show_version():
    messagebox.showinfo(
        "버전 정보",
        f"근태 자동정리 프로그램 {APP_VERSION}\n"
        f"빌드일자: {BUILD_DATE}\n"
        f"개발자: {DEVELOPER}"
    )

def run_app():
    root = tk.Tk()
    root.title("근태 자동정리 프로그램")
    root.geometry("500x350")

    tk.Label(root, text="근태 자동정리 프로그램", font=("맑은 고딕", 14, "bold")).grid(row=0, column=0, columnspan=3, pady=15)

    def select_office_file():
        path = filedialog.askopenfilename(title="사무실 근태 파일 선택", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            office_var.set(path)

    def select_factory_file():
        path = filedialog.askopenfilename(title="공장 근태 파일 선택", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            factory_var.set(path)

    def execute():
        # 년월 입력 확인
        year_str = year_entry.get().strip()
        month_str = month_entry.get().strip()
        if not (year_str and month_str):
            messagebox.showerror("오류", "년도와 월을 모두 입력해주세요.")
            return

        factory_path = factory_var.get().strip()
        office_path = office_var.get().strip()

        try:
            factory_df = process_factory(factory_path) if factory_path else None
            office_df = process_office(office_path) if office_path else None
            merged_df = merge_data(factory_df, office_df)
            out_file = update_excel(factory_df, office_df, merged_df, year_str, month_str)
            messagebox.showinfo("완료", f"근태 파일 갱신 완료!\n{out_file}")
        except Exception as e:
            import traceback
            messagebox.showerror("에러 발생", traceback.format_exc())

    # 폰트 및 정렬
    label_font = ("맑은 고딕", 11)
    entry_width = 25
    button_width = 15

    # UI 구성
    tk.Label(root, text="년도:", font=label_font).grid(row=1, column=0, padx=10, pady=10, sticky="e")
    year_entry = tk.Entry(root, width=entry_width)
    year_entry.grid(row=1, column=1, padx=5)

    tk.Label(root, text="월:", font=label_font).grid(row=2, column=0, padx=10, pady=10, sticky="e")
    month_entry = tk.Entry(root, width=entry_width)
    month_entry.grid(row=2, column=1, padx=5)

    office_var = tk.StringVar()
    factory_var = tk.StringVar()

    tk.Label(root, text="사무실 근태 파일:", font=label_font).grid(row=3, column=0, padx=10, pady=10, sticky="e")
    tk.Entry(root, textvariable=office_var, width=entry_width).grid(row=3, column=1)
    tk.Button(root, text="찾기", command=select_office_file, width=button_width).grid(row=3, column=2, padx=5)

    tk.Label(root, text="공장 근태 파일:", font=label_font).grid(row=4, column=0, padx=10, pady=10, sticky="e")
    tk.Entry(root, textvariable=factory_var, width=entry_width).grid(row=4, column=1)
    tk.Button(root, text="찾기", command=select_factory_file, width=button_width).grid(row=4, column=2, padx=5)

    tk.Button(root, text="실행", command=execute, width=button_width, bg="#4CAF50", fg="white").grid(row=5, column=1, pady=15)
    tk.Button(root, text="버전 정보", command=show_version, width=button_width).grid(row=6, column=1, pady=5)

    root.mainloop()

if __name__ == "__main__":
    run_app()
