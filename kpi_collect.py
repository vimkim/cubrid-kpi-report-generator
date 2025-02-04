#!/usr/bin/env python
import os
import glob
import argparse
import pandas as pd


def process_folder(folder_path):
    # 지정된 폴더 내의 모든 .xlsx 파일 목록 가져오기
    excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    if not excel_files:
        print(f"[INFO] 지정한 폴더({folder_path}) 내에 Excel 파일이 없습니다.")
        return

    # 결과를 저장할 빈 DataFrame 생성
    all_done_entries = pd.DataFrame()

    # 각 파일에서 'done' 상태의 행을 필터링하여 누적
    for file in excel_files:
        try:
            # 헤더가 없는 경우 header=None, 헤더가 있다면 옵션을 조정하세요.
            df = pd.read_excel(file, header=None, engine="openpyxl")
        except Exception as e:
            print(f"[ERROR] 파일 {file} 읽기 실패: {e}")
            continue

        # 예제에서는 4번째 열(인덱스 3)이 상태(status) 열이라고 가정
        # 문자열로 변환한 후 소문자로 비교
        done_rows = df[df[5].astype(str).str.strip().str.lower() == "done"]
        if not done_rows.empty:
            all_done_entries = pd.concat(
                [all_done_entries, done_rows], ignore_index=True
            )

    # 결과가 존재할 경우, 새로운 엑셀 파일로 저장
    if not all_done_entries.empty:
        output_file = os.path.join(folder_path, "done_entries.xlsx")
        try:
            all_done_entries.to_excel(output_file, index=False)
            print(f"[SUCCESS] 모든 'done' 항목이 '{output_file}'에 저장되었습니다.")
        except Exception as e:
            print(f"[ERROR] 결과 파일 저장 실패: {e}")
    else:
        print("[INFO] 'done' 항목이 발견되지 않았습니다.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="지정 폴더 내의 모든 xlsx 파일을 순회하여 'done' 상태의 항목을 모읍니다."
    )
    parser.add_argument(
        "folder", help="Excel 파일들이 위치한 폴더의 경로 (예: /path/to/folder)"
    )
    args = parser.parse_args()
    process_folder(args.folder)
