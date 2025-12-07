# WEBPRO入力シート統合スクリプト 使用方法

## 概要

100個のWEBPRO入力シート（Excel）を**1シート・295列**のデータベースに統合するスクリプトです。

## 出力仕様

| 項目 | 内容 |
|------|------|
| 行数 | 約10,000〜40,000行（100建物分） |
| 列数 | **295列** |
| キー列 | `file_id`, `entity_type` |
| 形式 | Excel (.xlsx) |

## 使用方法

### 基本コマンド

```bash
python consolidate_webpro_full.py --input_dir ./webpro_files --output ./webpro_all_data.xlsx
```

### オプション

| オプション | 説明 | デフォルト |
|------------|------|------------|
| `--input_dir`, `-i` | WEBPROファイルが格納されたディレクトリ | （必須） |
| `--output`, `-o` | 出力Excelファイルパス | `webpro_all_data.xlsx` |
| `--pattern`, `-p` | ファイルパターン | `*.xlsx` |

### 例

```bash
# 基本的な使用
python consolidate_webpro_full.py -i ./input_files -o ./output/webpro_database.xlsx

# 特定のファイル名パターンを指定
python consolidate_webpro_full.py -i ./input_files -o ./output.xlsx -p "WEBPRO_*.xlsx"
```

## 必要なライブラリ

```bash
pip install pandas openpyxl
```

## 出力データ構造

### entity_type（データ種別）一覧

| entity_type | 対応様式 | 内容 |
|-------------|----------|------|
| room | 様式1 | 室仕様 |
| zone | 様式2-1 | 空調ゾーン |
| wall | 様式2-2 | 外壁構成 |
| window | 様式2-3 | 窓仕様 |
| envelope | 様式2-4 | 外皮 |
| heatsource | 様式2-5 | 熱源 |
| pump | 様式2-6 | 二次ポンプ |
| ahu | 様式2-7 | 空調機 |
| hs_water_temp | 様式2-8 | 熱源水温度 |
| heat_exchanger | 様式2-9 | 全熱交換器 |
| vwv_pump | 様式2-10 | 変流量二次ポンプ |
| pac_partial | 様式2-11 | PAC部分負荷特性 |
| vent_room | 様式3-1 | 換気室 |
| vent_fan | 様式3-2 | 換気送風機 |
| vent_ahu | 様式3-3 | 換気空調機 |
| vent_load_rate | 様式3-4 | 年間平均負荷率 |
| lighting | 様式4 | 照明 |
| hotwater_room | 様式5-1 | 給湯室 |
| hotwater_equip | 様式5-2 | 給湯機器 |
| elevator | 様式6 | 昇降機 |
| pv | 様式7-1 | 太陽光発電 |
| cgs | 様式7-3 | コージェネ |
| envelope_non_ac | 様式8 | 非空調外皮 |

## Python活用例

```python
import pandas as pd

# データ読み込み
df = pd.read_excel('webpro_all_data.xlsx')

# 基本情報
print(f"総行数: {len(df)}")
print(f"建物数: {df['file_id'].nunique()}")

# entity_type別に抽出
df_rooms = df[df['entity_type'] == 'room']
df_hs = df[df['entity_type'] == 'heatsource']
df_light = df[df['entity_type'] == 'lighting']

# 建物ごとの延床面積
total_area = df_rooms.groupby('file_id')['room_area'].sum()

# 熱源種別の分布
hs_dist = df_hs['hs_type'].value_counts()

# 特定建物のデータ
building_001 = df[df['file_id'] == '001']

# BEI近似式用の集約（必要に応じて）
building_summary = df_rooms.groupby('file_id').agg({
    'building_name': 'first',
    'region': 'first',
    'room_area': 'sum'
}).reset_index()
```

## 列定義の詳細

全295列の詳細定義は `webpro_complete_column_definition.md` を参照してください。

## ファイル一覧

| ファイル | 説明 |
|----------|------|
| `consolidate_webpro_full.py` | 統合スクリプト本体 |
| `webpro_complete_column_definition.md` | 全295列の詳細定義 |
| `webpro_all_data.xlsx` | 出力ファイル（実行後生成） |

## 注意事項

1. **入力ファイル名**: ファイル名のソート順で`file_id`が割り当てられます（001〜100）
2. **文字コード**: 日本語を含むため、UTF-8環境での実行を推奨
3. **メモリ**: 100ファイル処理時は十分なメモリ（4GB以上推奨）を確保
4. **NULL値**: 該当しないデータ種別の列は空白（NULL）になります
