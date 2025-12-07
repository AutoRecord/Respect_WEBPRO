#!/usr/bin/env python3
"""
WEBPRO入力シート統合スクリプト（サンプル）
100個のWEBPRO入力シートを1つのExcelファイルに統合する
"""

import pandas as pd
from pathlib import Path
from typing import Dict, List, Tuple
import re

# ============================================
# 設定
# ============================================

# 統合対象シートの定義（シート名: 出力シート名）
SHEET_CONFIG = {
    '0) 基本情報': {'output_name': '00_基本情報', 'type': 'vertical'},
    '1) 室仕様': {'output_name': '01_室仕様', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 9, 'data_cols': 14},
    '2-1) 空調ゾーン': {'output_name': '02_空調ゾーン', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 9, 'data_cols': 12},
    '2-2) 外壁構成 ': {'output_name': '03_外壁構成', 'type': 'horizontal', 'header_row': 4, 'unit_row': 7, 'data_start': 9, 'data_cols': 9},
    '2-3) 窓仕様': {'output_name': '04_窓仕様', 'type': 'horizontal', 'header_row': 4, 'unit_row': 7, 'data_start': 9, 'data_cols': 8},
    '2-4) 外皮 ': {'output_name': '05_外皮', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 9, 'data_cols': 10},
    '2-5) 熱源': {'output_name': '06_熱源', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 9, 'data_cols': 24},
    '2-6) 2次ﾎﾟﾝﾌﾟ': {'output_name': '07_二次ポンプ', 'type': 'horizontal', 'header_row': 4, 'unit_row': 7, 'data_start': 9, 'data_cols': 10},
    '2-7) 空調機': {'output_name': '08_空調機', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 9, 'data_cols': 24},
    '2-9) 全熱交換器': {'output_name': '09_全熱交換器', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 9, 'data_cols': 18},
    '3-1) 換気室': {'output_name': '10_換気室', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 9, 'data_cols': 7},
    '3-2) 換気送風機': {'output_name': '11_換気送風機', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 9, 'data_cols': 6},
    '3-3) 換気空調機': {'output_name': '12_換気空調機', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 8, 'data_cols': 11},
    '4) 照明': {'output_name': '13_照明', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 9, 'data_cols': 17},
    '5-1) 給湯室': {'output_name': '14_給湯室', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 9, 'data_cols': 8},
    '5-2) 給湯機器': {'output_name': '15_給湯機器', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 9, 'data_cols': 9},
    '6) 昇降機': {'output_name': '16_昇降機', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 9, 'data_cols': 10},
    '7-1) 太陽光発電': {'output_name': '17_太陽光発電', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 9, 'data_cols': 7},
    '7-3) コージェネレーション設備': {'output_name': '18_コージェネ', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 8, 'data_cols': 16},
    '8) 非空調外皮': {'output_name': '19_非空調外皮', 'type': 'horizontal', 'header_row': 5, 'unit_row': 7, 'data_start': 9, 'data_cols': 14},
}


def extract_basic_info(df: pd.DataFrame, file_id: str) -> pd.DataFrame:
    """基本情報シート（縦型フォーム）からデータを抽出"""
    info = {'file_id': file_id}
    
    # 行ごとに項目を抽出
    key_map = {
        'シート作成月日': 'sheet_date',
        '入力責任者': 'responsible_person',
        '評価対象': 'evaluation_target',
        '建物の名称': 'building_name',
        '建築物所在地': 'location',
        '省エネ基準地域区分': 'region_class',
        '構造': 'structure',
        '階数': 'floor_count',
    }
    
    for idx, row in df.iterrows():
        for key, eng_key in key_map.items():
            if pd.notna(row.iloc[1]) and key in str(row.iloc[1]):
                # 値は通常C列以降
                value = row.iloc[2] if len(row) > 2 and pd.notna(row.iloc[2]) else ''
                info[eng_key] = value
    
    return pd.DataFrame([info])


def extract_horizontal_data(
    df: pd.DataFrame,
    file_id: str,
    building_name: str,
    config: dict
) -> pd.DataFrame:
    """横型テーブルからデータを抽出"""
    header_row = config['header_row']
    unit_row = config.get('unit_row')
    data_start = config['data_start']
    data_cols = config['data_cols']
    
    # ヘッダー取得
    headers = df.iloc[header_row, :data_cols].tolist()
    
    # 単位があれば結合
    if unit_row and unit_row < len(df):
        units = df.iloc[unit_row, :data_cols].tolist()
        headers = [
            f"{h}_{u}" if pd.notna(u) and u != '' and not str(u).startswith('(') else str(h) 
            for h, u in zip(headers, units)
        ]
    
    # ヘッダーをクリーンアップ
    headers = [str(h).strip().replace('\n', '') if pd.notna(h) else f'col_{i}' 
               for i, h in enumerate(headers)]
    
    # データ行を抽出
    data_df = df.iloc[data_start:, :data_cols].copy()
    data_df.columns = headers
    
    # 空白行を除去（最初の列が空白の行）
    first_col = headers[0]
    data_df = data_df[data_df[first_col].notna() & (data_df[first_col] != '')]
    
    # file_idとbuilding_nameを追加
    data_df.insert(0, 'file_id', file_id)
    data_df.insert(1, 'building_name', building_name)
    
    return data_df


def process_single_file(file_path: Path, file_id: str) -> Dict[str, pd.DataFrame]:
    """1つのファイルから全シートのデータを抽出"""
    results = {}
    xl = pd.ExcelFile(file_path)
    
    # まず基本情報から建物名を取得
    building_name = ''
    if '0) 基本情報' in xl.sheet_names:
        df_info = pd.read_excel(file_path, sheet_name='0) 基本情報', header=None)
        info_df = extract_basic_info(df_info, file_id)
        results['00_基本情報'] = info_df
        building_name = info_df.iloc[0].get('building_name', '')
    
    # 各シートを処理
    for sheet_name, config in SHEET_CONFIG.items():
        if sheet_name == '0) 基本情報':
            continue  # 既に処理済み
        
        if sheet_name not in xl.sheet_names:
            continue
        
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            
            if config['type'] == 'horizontal':
                extracted = extract_horizontal_data(df, file_id, building_name, config)
                if len(extracted) > 0:
                    results[config['output_name']] = extracted
        except Exception as e:
            print(f"  警告: {sheet_name} の処理中にエラー: {e}")
    
    return results


def consolidate_files(input_dir: Path, output_path: Path):
    """複数のWEBPROファイルを統合"""
    
    # 入力ファイルを取得
    input_files = sorted(input_dir.glob('*.xlsx'))
    print(f"入力ファイル数: {len(input_files)}")
    
    # 結果を格納する辞書
    all_data: Dict[str, List[pd.DataFrame]] = {}
    
    # 各ファイルを処理
    for i, file_path in enumerate(input_files, 1):
        file_id = f"{i:03d}"  # 001, 002, ...
        print(f"処理中: [{file_id}] {file_path.name}")
        
        try:
            file_results = process_single_file(file_path, file_id)
            
            for sheet_name, df in file_results.items():
                if sheet_name not in all_data:
                    all_data[sheet_name] = []
                all_data[sheet_name].append(df)
        except Exception as e:
            print(f"  エラー: {file_path.name} の処理に失敗: {e}")
    
    # データを結合して出力
    print(f"\n統合ファイルを出力中: {output_path}")
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for sheet_name in sorted(all_data.keys()):
            combined_df = pd.concat(all_data[sheet_name], ignore_index=True)
            combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  {sheet_name}: {len(combined_df)}行")
    
    print("\n統合完了！")


# ============================================
# メイン処理
# ============================================

if __name__ == '__main__':
    # 使用例
    # python consolidate_webpro.py
    
    # パスを設定（実際の使用時は変更）
    input_directory = Path('./input_files')  # 100個のファイルを配置
    output_file = Path('./output/webpro_combined_data.xlsx')
    
    # 出力ディレクトリを作成
    output_file.parent.mkdir(parents=True, exist_ok=True)
    
    # 統合実行
    consolidate_files(input_directory, output_file)
