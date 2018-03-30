# -*- coding: utf-8 -*-
import json
import pandas as pd
import numpy as np
import gspread
from oauth2client.service_account import ServiceAccountCredentials

class GspreadHandler(object):
    """
        python から Google SpreadSheet を使うときの関数など集めたクラス
    """
    def __init__(self, json_file_path):
        self.json_file = json_file_path
        self.scope = ['https://spreadsheets.google.com/feeds']
        self.gsfile = None


    def get_gsfile(self, sheet_id):
        """ 読み書きするスプレッドシートのインスタンスを取得 """
        credentials = ServiceAccountCredentials.from_json_keyfile_name(self.json_file, self.scope)
        client = gspread.authorize(credentials)
        self.gsfile = client.open_by_key(sheet_id)


    def dataset_fromSheet(self, read_sheet_name):
        """
            スプレッドシートの該当シートを読み込んで前処理を行う。出力は DataFrame 型。

            read_sheet_name: 読み込み元シート名
        """

        # データ読み込み
        raw_data_sheet = self.gsfile.worksheet(read_sheet_name)
        dataset = pd.DataFrame(raw_data_sheet.get_all_values())

        # 0行目を列名にする
        dataset.columns =  list(dataset.iloc[0])
        dataset = dataset.drop(0, axis=0)
        dataset = dataset.reset_index(drop=True)

        return dataset


    def update_CellsOfColumn(self, dataset_series, write_sheet_name, column_alphabet):
        """
            スプレッドシートのある列を一気に更新する関数。
            ただし1行目には既にヘッダが入力されている想定で2行目以降に数値を入れる。

            dataset_series: 更新用dataframeの該当列。
            update_sheet: アップデート先スプレッドシートのインスタンス。
            column_alphabet: 更新する列名のアルファベット。
        """

        update_sheet = self.gsfile.worksheet(write_sheet_name)

        # 更新範囲指定。A2:A100 のような文字列をつくる
        data_range = column_alphabet + '2:' + column_alphabet + str(len(dataset_series)+1)
        cell_list = update_sheet.range(data_range)

        for (i, cell) in enumerate(cell_list):
            cell.value = str(dataset_series[i])

        update_sheet.update_cells(cell_list)
