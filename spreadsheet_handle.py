# -*- coding: utf-8 -*-
import pandas as pd
from gshandler import GspreadHandler

json_file = 'YOUR_JSON_FILE_PATH'

# URL https://docs.google.com/spreadsheets/d/xxxxxxxxxxx の xxxxxxxxxxx 部分のこと
sheet_id = 'SPREADSHEET_ID'

read_sheet_name = 'dataset'
write_sheet_name = 'reporting'

def main():
    # スプレッドシートからデータを収集する。

    # スプレッドシートのクラスのインスタンスを作成、ファイル情報を読みとる
    gh = GspreadHandler(json_file)
    gh.get_gsfile(sheet_id)

    # read_sheet_name のシートの値を取得。DataFrame型で取得。
    dataset = gh.dataset_fromSheet(read_sheet_name)

    # spreadsheet から読み取れるデータは文字列型なので、floatに直す。
    dataset['hoge_data'] = dataset['hoge_data'].astype(float)
    dataset['fuga_data'] = dataset['fuga_data'].astype(float)

    # KPIを計算する。
    dataset['hogehoge_KPI'] = dataset['hoge_data'] + dataset['fuga_data']

    # write_sheet_name のシートにデータを書き込む。
    # 1列ごとにまとめて書き込む。書き込み先の列名はアルファベットで指定。
    gh.update_CellsOfColumn(dataset['id'], write_sheet_name, "A")
    gh.update_CellsOfColumn(dataset['hoge_data'], write_sheet_name, "B")
    gh.update_CellsOfColumn(dataset['fuga_data'], write_sheet_name, "C")
    gh.update_CellsOfColumn(dataset['hogehoge_KPI'], write_sheet_name, "D")


if __name__ == '__main__':
    main()
