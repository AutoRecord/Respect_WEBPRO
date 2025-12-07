#!/usr/bin/env python3
"""
WEBPRO統合データ読み込みサンプル
統合後のExcelファイルをPythonで読み込んで分析する例
"""

import pandas as pd
from pathlib import Path


# ============================================
# 読み込みパターン
# ============================================

def load_all_sheets(file_path: str) -> dict:
    """全シートを辞書形式で読み込み"""
    return pd.read_excel(file_path, sheet_name=None)


def load_specific_sheets(file_path: str, sheets: list) -> dict:
    """指定したシートのみ読み込み"""
    return pd.read_excel(file_path, sheet_name=sheets)


def load_single_sheet(file_path: str, sheet_name: str) -> pd.DataFrame:
    """1シートだけ読み込み"""
    return pd.read_excel(file_path, sheet_name=sheet_name)


# ============================================
# 分析サンプル
# ============================================

def example_analysis(file_path: str):
    """分析例"""
    
    print("=" * 60)
    print("WEBPRO統合データ分析サンプル")
    print("=" * 60)
    
    # 全シート読み込み
    all_data = load_all_sheets(file_path)
    
    print(f"\n【読み込んだシート一覧】")
    for name, df in all_data.items():
        print(f"  {name}: {len(df)}行 × {len(df.columns)}列")
    
    # ------------------------------
    # サンプル1: 建物一覧
    # ------------------------------
    if '00_基本情報' in all_data:
        df_info = all_data['00_基本情報']
        print(f"\n【建物一覧】")
        print(df_info[['file_id', 'building_name', 'location']].head(10))
    
    # ------------------------------
    # サンプル2: 室面積の集計
    # ------------------------------
    if '01_室仕様' in all_data:
        df_rooms = all_data['01_室仕様']
        
        # 数値変換（文字列の場合に備えて）
        df_rooms['室面積'] = pd.to_numeric(df_rooms['室面積'], errors='coerce')
        
        print(f"\n【建物別 室面積合計】")
        room_summary = df_rooms.groupby('file_id').agg({
            'building_name': 'first',
            '室面積': 'sum',
            '室名': 'count'
        }).rename(columns={'室名': '室数'})
        print(room_summary.head(10))
    
    # ------------------------------
    # サンプル3: 熱源機器の集計
    # ------------------------------
    if '06_熱源' in all_data:
        df_heat = all_data['06_熱源']
        
        print(f"\n【熱源機種の使用状況】")
        if '熱源機種' in df_heat.columns:
            heat_type_counts = df_heat['熱源機種'].value_counts()
            print(heat_type_counts.head(10))
    
    # ------------------------------
    # サンプル4: 照明制御の採用状況
    # ------------------------------
    if '13_照明' in all_data:
        df_light = all_data['13_照明']
        
        print(f"\n【照明制御の採用状況】")
        control_cols = [c for c in df_light.columns if '制御' in c]
        for col in control_cols:
            print(f"\n  {col}:")
            print(df_light[col].value_counts())
    
    # ------------------------------
    # サンプル5: 特定建物の詳細抽出
    # ------------------------------
    print(f"\n【特定建物（file_id=001）のデータ抽出】")
    for sheet_name, df in all_data.items():
        if 'file_id' in df.columns:
            building_data = df[df['file_id'] == '001']
            if len(building_data) > 0:
                print(f"  {sheet_name}: {len(building_data)}行")


# ============================================
# CSVエクスポート
# ============================================

def export_to_csv(file_path: str, output_dir: str):
    """各シートをCSVとして出力（Python以外のツール連携用）"""
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    
    all_data = load_all_sheets(file_path)
    
    for name, df in all_data.items():
        csv_path = output_path / f"{name}.csv"
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        print(f"出力: {csv_path}")


# ============================================
# クエリ用ユーティリティ
# ============================================

class WebproData:
    """WEBPRO統合データへの便利なアクセスを提供"""
    
    def __init__(self, file_path: str):
        self.file_path = file_path
        self._cache = {}
    
    def get_sheet(self, sheet_name: str) -> pd.DataFrame:
        """シートを取得（キャッシュ付き）"""
        if sheet_name not in self._cache:
            self._cache[sheet_name] = pd.read_excel(
                self.file_path, sheet_name=sheet_name
            )
        return self._cache[sheet_name]
    
    def get_building(self, file_id: str, sheet_name: str) -> pd.DataFrame:
        """特定建物の特定シートデータを取得"""
        df = self.get_sheet(sheet_name)
        return df[df['file_id'] == file_id]
    
    def get_all_buildings(self) -> pd.DataFrame:
        """全建物の基本情報を取得"""
        return self.get_sheet('00_基本情報')
    
    def search_rooms(self, room_type: str = None, min_area: float = None) -> pd.DataFrame:
        """室を検索"""
        df = self.get_sheet('01_室仕様')
        
        if room_type:
            df = df[df['室用途_小分類'].str.contains(room_type, na=False)]
        
        if min_area:
            df['室面積'] = pd.to_numeric(df['室面積'], errors='coerce')
            df = df[df['室面積'] >= min_area]
        
        return df


# ============================================
# メイン
# ============================================

if __name__ == '__main__':
    # 統合ファイルのパス
    combined_file = './output/webpro_combined_data.xlsx'
    
    # 分析実行
    example_analysis(combined_file)
    
    # CSVエクスポート（必要に応じて）
    # export_to_csv(combined_file, './output/csv/')
    
    # クラスを使った例
    # data = WebproData(combined_file)
    # print(data.get_building('001', '01_室仕様'))
    # print(data.search_rooms(room_type='事務室', min_area=100))
