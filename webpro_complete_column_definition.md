# WEBPRO全データ統合 完全列定義書

## 概要

- **方式**: 縦展開（1レコード=1行）
- **行数**: 約10,000〜40,000行（100建物分）
- **列数**: 約200列
- **キー列**: `file_id`, `entity_type`

---

## 列定義一覧

### A. 共通列（全行に値あり）

| # | 列名 | 説明 | 型 | 備考 |
|---|------|------|-----|------|
| 1 | file_id | ファイル識別子 | str | 001〜100 |
| 2 | building_name | 建物の名称 | str | 様式0より |
| 3 | prefecture | 都道府県 | str | 様式0より |
| 4 | city | 市区町村 | str | 様式0より |
| 5 | region | 省エネ基準地域区分 | int | 1〜8 |
| 6 | structure | 構造 | str | RC/S/SRC等 |
| 7 | floors_above | 地上階数 | int | |
| 8 | floors_below | 地下階数 | int | |
| 9 | evaluation_target | 評価対象 | str | 様式0より |
| 10 | **entity_type** | データ種別 | str | 下記参照 |

#### entity_type の値一覧

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

---

### B. 様式1：室仕様（entity_type = 'room'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 11 | room_floor | 階 | str | |
| 12 | room_name | 室名 | str | |
| 13 | room_building_type | 建物用途 | str | 選択 |
| 14 | room_type_major | 室用途（大分類） | str | 選択 |
| 15 | room_type_minor | 室用途（小分類） | str | 選択 |
| 16 | room_area | 室面積 | float | ㎡ |
| 17 | room_floor_height | 階高 | float | m |
| 18 | room_ceiling_height | 天井高 | float | m |
| 19 | room_is_ac_target | 空調計算対象室 | str | 選択 |
| 20 | room_is_vent_target | 換気計算対象室 | str | 選択 |
| 21 | room_is_light_target | 照明計算対象室 | str | 選択 |
| 22 | room_is_hotwater_target | 給湯計算対象室 | str | 選択 |
| 23 | room_building_group | 建築物の名称 | str | 複数建築物用 |
| 24 | room_note | 備考 | str | |

---

### C. 様式2-1：空調ゾーン（entity_type = 'zone'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 25 | zone_floor | 階 | str | 転記 |
| 26 | zone_room_name | 室名 | str | 転記 |
| 27 | zone_room_type_major | 室用途（大分類） | str | 転記 |
| 28 | zone_room_type_minor | 室用途（小分類） | str | 転記 |
| 29 | zone_room_area | 室面積 | float | ㎡ |
| 30 | zone_floor_height | 階高 | float | m |
| 31 | zone_ceiling_height | 天井高 | float | m |
| 32 | zone_ac_floor | 空調ゾーン階 | str | |
| 33 | zone_name | 空調ゾーン名 | str | |
| 34 | zone_ahu_group_room | 室負荷処理（空調機群名称） | str | 転記 |
| 35 | zone_ahu_group_oa | 外気負荷処理（空調機群名称） | str | 転記 |
| 36 | zone_note | 備考 | str | |

---

### D. 様式2-2：外壁構成（entity_type = 'wall'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 37 | wall_name | 外壁名称 | str | |
| 38 | wall_type | 壁の種類 | str | 選択（外壁/接地壁） |
| 39 | wall_u_value | 熱貫流率 | float | W/㎡K |
| 40 | wall_material_no | 建材番号 | int | 選択 |
| 41 | wall_material_name | 建材名称 | str | 選択 |
| 42 | wall_conductivity | 熱伝導率 | float | W/mK |
| 43 | wall_thickness | 厚み | float | mm |
| 44 | wall_solar_absorption | 日射吸収率 | float | - |
| 45 | wall_note | 備考 | str | |

---

### E. 様式2-3：窓仕様（entity_type = 'window'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 46 | window_name | 開口部名称 | str | |
| 47 | window_u_value | 窓の熱貫流率 | float | W/㎡K |
| 48 | window_eta_value | 窓の日射熱取得率 | float | - |
| 49 | window_frame_type | 建具の種類 | str | 選択 |
| 50 | window_glass_type | ガラスの種類 | str | 選択 |
| 51 | window_glass_u_value | ガラスの熱貫流率 | float | W/(㎡･K) |
| 52 | window_glass_eta_value | ガラスの日射熱取得率 | float | - |
| 53 | window_note | 備考 | str | |

---

### F. 様式2-4：外皮（entity_type = 'envelope'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 54 | env_floor | 階 | str | 転記 |
| 55 | env_zone_name | 空調ゾーン名 | str | 転記 |
| 56 | env_direction | 方位 | str | 選択 |
| 57 | env_shade_coef_cooling | 日除け効果係数（冷房） | float | - |
| 58 | env_shade_coef_heating | 日除け効果係数（暖房） | float | - |
| 59 | env_wall_name | 外壁名称 | str | 転記 |
| 60 | env_wall_area | 外皮面積（窓含） | float | ㎡ |
| 61 | env_window_name | 開口部名称 | str | 転記 |
| 62 | env_window_area | 窓面積 | float | ㎡ |
| 63 | env_has_blind | ブラインドの有無 | str | 選択 |
| 64 | env_note | 備考 | str | |

---

### G. 様式2-5：熱源（entity_type = 'heatsource'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 65 | hs_group_name | 熱源群名称 | str | |
| 66 | hs_simultaneous | 冷暖同時供給有無 | str | 選択 |
| 67 | hs_staging_control | 台数制御 | str | 選択 |
| 68 | hs_operation_mode | 運転モード | str | 選択（無/追掛/氷蓄熱/水蓄熱） |
| 69 | hs_storage_capacity | 蓄熱容量 | float | MJ |
| 70 | hs_type | 熱源機種 | str | 選択 |
| 71 | hs_cooling_order | 冷熱生成運転順位 | int | 選択 |
| 72 | hs_cooling_count | 冷熱生成台数 | int | 台 |
| 73 | hs_cooling_supply_temp | 冷熱生成送水温度 | float | ℃ |
| 74 | hs_cooling_capacity | 定格冷却能力 | float | kW/台 |
| 75 | hs_cooling_main_power | 主機定格消費エネルギー（冷） | float | kW/台 |
| 76 | hs_cooling_sub_power | 補機定格消費電力（冷） | float | kW/台 |
| 77 | hs_cooling_pump_power | 一次ポンプ定格消費電力（冷） | float | kW/台 |
| 78 | hs_ct_capacity | 冷却塔定格冷却能力 | float | kW/台 |
| 79 | hs_ct_fan_power | 冷却塔ファン消費電力 | float | kW/台 |
| 80 | hs_ct_pump_power | 冷却水ポンプ消費電力 | float | kW/台 |
| 81 | hs_heating_order | 温熱生成運転順位 | int | 選択 |
| 82 | hs_heating_count | 温熱生成台数 | int | 台 |
| 83 | hs_heating_supply_temp | 温熱生成送水温度 | float | ℃ |
| 84 | hs_heating_capacity | 定格加熱能力 | float | kW/台 |
| 85 | hs_heating_main_power | 主機定格消費エネルギー（温） | float | kW/台 |
| 86 | hs_heating_sub_power | 補機定格消費電力（温） | float | kW/台 |
| 87 | hs_heating_pump_power | 一次ポンプ定格消費電力（温） | float | kW/台 |
| 88 | hs_note | 備考 | str | |

---

### H. 様式2-6：二次ポンプ（entity_type = 'pump'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 89 | pump_group_name | 二次ポンプ群名称 | str | |
| 90 | pump_staging_control | 台数制御の有無 | str | 選択 |
| 91 | pump_cooling_temp_diff | 冷房時温度差 | float | ℃ |
| 92 | pump_heating_temp_diff | 暖房時温度差 | float | ℃ |
| 93 | pump_order | 運転順位 | int | 選択 |
| 94 | pump_count | 台数 | int | 台 |
| 95 | pump_rated_flow | 定格流量 | float | m3/h台 |
| 96 | pump_rated_power | 定格消費電力 | float | kW/台 |
| 97 | pump_flow_control | 流量制御方式 | str | 選択 |
| 98 | pump_min_flow_ratio | 変流量時最小流量比 | float | % |
| 99 | pump_note | 備考 | str | |

---

### I. 様式2-7：空調機（entity_type = 'ahu'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 100 | ahu_group_name | 空調機群名称 | str | |
| 101 | ahu_count | 台数 | int | 台 |
| 102 | ahu_type | 空調機タイプ | str | 選択 |
| 103 | ahu_cooling_capacity | 定格冷却（冷房）能力 | float | kW/台 |
| 104 | ahu_heating_capacity | 定格加熱（暖房）能力 | float | kW/台 |
| 105 | ahu_oa_flow | 設計最大外気風量 | float | m3/h台 |
| 106 | ahu_sa_fan_power | 給気ファン定格消費電力 | float | kW/台 |
| 107 | ahu_ra_fan_power | 還気ファン定格消費電力 | float | kW/台 |
| 108 | ahu_oa_fan_power | 外気ファン定格消費電力 | float | kW/台 |
| 109 | ahu_ea_fan_power | 排気ファン定格消費電力 | float | kW/台 |
| 110 | ahu_air_flow_control | 風量制御方式 | str | 選択 |
| 111 | ahu_min_air_ratio | 変風量時最小風量比 | float | % |
| 112 | ahu_preheat_oa_stop | 予熱時外気取り入れ停止の有無 | str | 選択 |
| 113 | ahu_economizer | 外気冷房制御の有無 | str | 選択 |
| 114 | ahu_has_hex | 全熱交換器の有無 | str | 選択 |
| 115 | ahu_hex_name | 全熱交換器の名称 | str | 転記 |
| 116 | ahu_hex_flow | 全熱交換器の設計風量 | float | m3/h台 |
| 117 | ahu_hex_eff_cooling | 全熱交換効率（冷房時） | float | % |
| 118 | ahu_hex_eff_heating | 全熱交換効率（暖房時） | float | % |
| 119 | ahu_auto_bypass | 自動換気切替機能の有無 | str | 選択 |
| 120 | ahu_rotor_power | ローター消費電力 | float | kW/台 |
| 121 | ahu_pump_group_cooling | 二次ポンプ群名称（冷熱） | str | 転記 |
| 122 | ahu_pump_group_heating | 二次ポンプ群名称（温熱） | str | 転記 |
| 123 | ahu_hs_group_cooling | 熱源群名称（冷熱） | str | 転記 |
| 124 | ahu_hs_group_heating | 熱源群名称（温熱） | str | 転記 |
| 125 | ahu_note | 備考 | str | |

---

### J. 様式2-8：熱源水温度（entity_type = 'hs_water_temp'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 126 | hswt_group_name | 熱源群名称 | str | |
| 127 | hswt_temp_jan | 熱源水温度（1月） | float | ℃ |
| 128 | hswt_temp_feb | 熱源水温度（2月） | float | ℃ |
| 129 | hswt_temp_mar | 熱源水温度（3月） | float | ℃ |
| 130 | hswt_temp_apr | 熱源水温度（4月） | float | ℃ |
| 131 | hswt_temp_may | 熱源水温度（5月） | float | ℃ |
| 132 | hswt_temp_jun | 熱源水温度（6月） | float | ℃ |
| 133 | hswt_temp_jul | 熱源水温度（7月） | float | ℃ |
| 134 | hswt_temp_aug | 熱源水温度（8月） | float | ℃ |
| 135 | hswt_temp_sep | 熱源水温度（9月） | float | ℃ |
| 136 | hswt_temp_oct | 熱源水温度（10月） | float | ℃ |
| 137 | hswt_temp_nov | 熱源水温度（11月） | float | ℃ |
| 138 | hswt_temp_dec | 熱源水温度（12月） | float | ℃ |

---

### K. 様式2-9：全熱交換器（entity_type = 'heat_exchanger'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 139 | hex_name | 全熱交換器名称 | str | |
| 140 | hex_type | 全熱交換器の方式 | str | 選択（静止形/回転形） |
| 141 | hex_oa_flow | 設計外気量又は設計給気量 | float | m3/h台 |
| 142 | hex_ea_flow | 設計排気量又は設計還気量 | float | m3/h台 |
| 143 | hex_count | 台数 | int | 台 |
| 144 | hex_eff_cooling_1 | 全熱交換効率（冷房時）試験1 | float | % |
| 145 | hex_eff_heating_1 | 全熱交換効率（暖房時）試験1 | float | % |
| 146 | hex_test_sa_flow_1 | 測定時給気量 試験1 | float | m3/h台 |
| 147 | hex_test_ra_flow_1 | 測定時還気量 試験1 | float | m3/h台 |
| 148 | hex_vent_eff_1 | 有効換気量率 試験1 | float | % |
| 149 | hex_eff_cooling_2 | 全熱交換効率（冷房時）試験2 | float | % |
| 150 | hex_eff_heating_2 | 全熱交換効率（暖房時）試験2 | float | % |
| 151 | hex_test_sa_flow_2 | 測定時給気量 試験2 | float | m3/h台 |
| 152 | hex_test_ra_flow_2 | 測定時還気量 試験2 | float | m3/h台 |
| 153 | hex_vent_eff_2 | 有効換気量率 試験2 | float | % |
| 154 | hex_eff_cooling_3 | 全熱交換効率（冷房時）試験3 | float | % |
| 155 | hex_eff_heating_3 | 全熱交換効率（暖房時）試験3 | float | % |
| 156 | hex_test_sa_flow_3 | 測定時給気量 試験3 | float | m3/h台 |
| 157 | hex_test_ra_flow_3 | 測定時還気量 試験3 | float | m3/h台 |
| 158 | hex_vent_eff_3 | 有効換気量率 試験3 | float | % |

---

### L. 様式2-10：変流量二次ポンプシステム（entity_type = 'vwv_pump'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 159 | vwv_group_name | 二次ポンプ群名称 | str | |
| 160 | vwv_cooling_temp_diff | 冷房時温度差 | float | ℃ |
| 161 | vwv_heating_temp_diff | 暖房時温度差 | float | ℃ |
| 162 | vwv_rated_flow | 定格流量 | float | m3/h |
| 163 | vwv_rated_power | 定格消費電力 | float | kW |
| 164 | vwv_min_flow_ratio | 変流量時最小流量比 | float | % |
| 165 | vwv_coef_3rd | 3次の項の係数 | float | |
| 166 | vwv_coef_2nd | 2次の項の係数 | float | |
| 167 | vwv_coef_1st | 1次の項の係数 | float | |
| 168 | vwv_coef_const | 定数項 | float | |

---

### M. 様式2-11：PAC部分負荷特性（entity_type = 'pac_partial'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 169 | pac_hs_name | 熱源機種名称 | str | |
| 170 | pac_cooling_coef_2nd | 部分負荷特性（冷房）2次の項の係数 | float | |
| 171 | pac_cooling_coef_1st | 部分負荷特性（冷房）1次の項の係数 | float | |
| 172 | pac_cooling_const | 部分負荷特性（冷房）定数項 | float | |
| 173 | pac_cooling_min_output | 部分負荷特性（冷房）最小出力比 | float | |
| 174 | pac_heating_coef_2nd | 部分負荷特性（暖房）2次の項の係数 | float | |
| 175 | pac_heating_coef_1st | 部分負荷特性（暖房）1次の項の係数 | float | |
| 176 | pac_heating_const | 部分負荷特性（暖房）定数項 | float | |
| 177 | pac_heating_min_output | 部分負荷特性（暖房）最小出力比 | float | |

---

### N. 様式3-1：換気室（entity_type = 'vent_room'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 178 | vr_floor | 階 | str | 転記 |
| 179 | vr_room_name | 室名 | str | 転記 |
| 180 | vr_room_type_major | 室用途（大分類） | str | 転記 |
| 181 | vr_room_type_minor | 室用途（小分類） | str | 転記 |
| 182 | vr_room_area | 室面積 | float | ㎡ |
| 183 | vr_vent_type | 換気種類 | str | 選択（給気/排気/循環/空調） |
| 184 | vr_vent_equip_name | 換気機器名称 | str | 転記 |
| 185 | vr_note | 備考 | str | |

---

### O. 様式3-2：換気送風機（entity_type = 'vent_fan'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 186 | vf_equip_name | 換気機器名称 | str | |
| 187 | vf_design_flow | 設計風量 | float | m3/h |
| 188 | vf_motor_power | 電動機定格出力 | float | kW |
| 189 | vf_high_eff_motor | 高効率電動機の有無 | str | 選択 |
| 190 | vf_has_inverter | インバータの有無 | str | 選択 |
| 191 | vf_flow_control | 送風量制御 | str | 選択 |
| 192 | vf_note | 備考 | str | |

---

### P. 様式3-3：換気空調機（entity_type = 'vent_ahu'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 193 | va_equip_name | 換気機器名称 | str | |
| 194 | va_room_type | 換気対象室の用途 | str | 選択 |
| 195 | va_cooling_capacity | 必要冷却能力 | float | kW |
| 196 | va_hs_efficiency | 熱源効率（一次換算値） | float | - |
| 197 | va_pump_power | ポンプ定格出力 | float | kW |
| 198 | va_fan_type | 送風機の種類 | str | 選択 |
| 199 | va_design_flow | 設計風量 | float | m3/h |
| 200 | va_motor_power | 電動機定格出力 | float | kW |
| 201 | va_high_eff_motor | 高効率電動機の有無 | str | 選択 |
| 202 | va_has_inverter | インバータの有無 | str | 選択 |
| 203 | va_flow_control | 送風量制御 | str | 選択 |
| 204 | va_note | 備考 | str | |

---

### Q. 様式3-4：年間平均負荷率（entity_type = 'vent_load_rate'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 205 | vlr_equip_name | 換気機器名称 | str | |
| 206 | vlr_annual_load_rate | 年間平均負荷率 | float | - |
| 207 | vlr_note | 備考 | str | |

---

### R. 様式4：照明（entity_type = 'lighting'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 208 | lt_floor | 階 | str | 転記 |
| 209 | lt_room_name | 室名 | str | 転記 |
| 210 | lt_room_type_major | 室用途（大分類） | str | 転記 |
| 211 | lt_room_type_minor | 室用途（小分類） | str | 転記 |
| 212 | lt_room_area | 室面積 | float | ㎡ |
| 213 | lt_floor_height | 階高 | float | m |
| 214 | lt_ceiling_height | 天井高 | float | m |
| 215 | lt_room_width | 室の間口 | float | m |
| 216 | lt_room_depth | 室の奥行 | float | m |
| 217 | lt_room_index | 室指数 | float | - |
| 218 | lt_fixture_name | 機器名称 | str | |
| 219 | lt_fixture_power | 定格消費電力 | float | W/台 |
| 220 | lt_fixture_count | 台数 | int | 台 |
| 221 | lt_occupancy_control | 在室検知制御 | str | 選択 |
| 222 | lt_daylight_control | 明るさ検知制御 | str | 選択 |
| 223 | lt_schedule_control | タイムスケジュール制御 | str | 選択 |
| 224 | lt_initial_correction | 初期照度補正機能 | str | 選択 |
| 225 | lt_note | 備考 | str | |

---

### S. 様式5-1：給湯室（entity_type = 'hotwater_room'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 226 | hwr_floor | 階 | str | 転記 |
| 227 | hwr_room_name | 室名 | str | 転記 |
| 228 | hwr_room_type_major | 室用途（大分類） | str | 転記 |
| 229 | hwr_room_type_minor | 室用途（小分類） | str | 転記 |
| 230 | hwr_room_area | 室面積 | float | ㎡ |
| 231 | hwr_supply_location | 給湯箇所 | str | |
| 232 | hwr_water_saving | 節湯器具 | str | 選択 |
| 233 | hwr_equip_name | 給湯機器名称 | str | 転記 |
| 234 | hwr_note | 備考 | str | |

---

### T. 様式5-2：給湯機器（entity_type = 'hotwater_equip'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 235 | hwe_equip_name | 給湯機器名称 | str | |
| 236 | hwe_fuel_type | 燃料種類 | str | 選択 |
| 237 | hwe_heating_capacity | 定格加熱能力 | float | kW |
| 238 | hwe_efficiency | 熱源効率（一次エネルギー換算） | float | - |
| 239 | hwe_insulation | 配管保温仕様 | str | 選択 |
| 240 | hwe_pipe_diameter | 接続口径 | float | mm |
| 241 | hwe_solar_area | 有効集熱面積 | float | ㎡ |
| 242 | hwe_solar_azimuth | 集熱面の方位角 | float | ° |
| 243 | hwe_solar_tilt | 集熱面の傾斜角 | float | ° |
| 244 | hwe_note | 備考 | str | |

---

### U. 様式6：昇降機（entity_type = 'elevator'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 245 | ev_floor | 階 | str | 転記 |
| 246 | ev_room_name | 室名 | str | 転記 |
| 247 | ev_room_type_major | 室用途（大分類） | str | 転記 |
| 248 | ev_room_type_minor | 室用途（小分類） | str | 転記 |
| 249 | ev_equip_name | 機器名称 | str | |
| 250 | ev_count | 台数 | int | 台 |
| 251 | ev_capacity | 積載量 | float | kg |
| 252 | ev_speed | 速度 | float | m/min |
| 253 | ev_transport_coef | 輸送能力係数 | float | - |
| 254 | ev_control_type | 速度制御方式 | str | 選択 |
| 255 | ev_note | 備考 | str | |

---

### V. 様式7-1：太陽光発電（entity_type = 'pv'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 256 | pv_system_name | 太陽光発電システム名称 | str | |
| 257 | pv_pcs_efficiency | パワーコンディショナの効率 | float | - |
| 258 | pv_cell_type | 太陽電池の種類 | str | 選択 |
| 259 | pv_install_type | アレイ設置方式 | str | 選択 |
| 260 | pv_capacity | アレイのシステム容量 | float | kW |
| 261 | pv_azimuth | パネルの方位角 | float | ° |
| 262 | pv_tilt | パネルの傾斜角 | float | ° |
| 263 | pv_note | 備考 | str | |

---

### W. 様式7-3：コージェネレーション（entity_type = 'cgs'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 264 | cgs_name | コージェネレーション設備名称 | str | |
| 265 | cgs_rated_output | 定格発電出力 | float | kW |
| 266 | cgs_count | 設置台数 | int | 台 |
| 267 | cgs_gen_eff_100 | 発電効率（負荷率1.00） | float | - |
| 268 | cgs_gen_eff_75 | 発電効率（負荷率0.75） | float | - |
| 269 | cgs_gen_eff_50 | 発電効率（負荷率0.50） | float | - |
| 270 | cgs_heat_eff_100 | 排熱効率（負荷率1.00） | float | - |
| 271 | cgs_heat_eff_75 | 排熱効率（負荷率0.75） | float | - |
| 272 | cgs_heat_eff_50 | 排熱効率（負荷率0.50） | float | - |
| 273 | cgs_priority_ac_cool | 排熱利用優先順位（空調冷熱源） | int | - |
| 274 | cgs_priority_ac_heat | 排熱利用優先順位（空調温熱源） | int | - |
| 275 | cgs_priority_hotwater | 排熱利用優先順位（給湯） | int | - |
| 276 | cgs_24h_operation | 24時間運転の有無 | str | - |
| 277 | cgs_ac_cool_hs_group | 排熱利用系統（空調冷熱源） | str | 選択 |
| 278 | cgs_ac_heat_hs_group | 排熱利用系統（空調温熱源） | str | 選択 |
| 279 | cgs_hotwater_equip | 排熱利用系統（給湯機器） | str | 選択 |
| 280 | cgs_note | 備考 | str | |

---

### X. 様式8：非空調外皮（entity_type = 'envelope_non_ac'）

| # | 列名 | 説明 | 型 | 単位 |
|---|------|------|-----|------|
| 281 | nac_floor | 階 | str | |
| 282 | nac_zone_name | 非空調ゾーン名 | str | |
| 283 | nac_room_type_major | 室用途（大分類） | str | |
| 284 | nac_room_type_minor | 室用途（小分類） | str | |
| 285 | nac_room_area | 室面積 | float | ㎡ |
| 286 | nac_floor_height | 階高 | float | m |
| 287 | nac_direction | 方位 | str | 選択 |
| 288 | nac_shade_coef_cooling | 日除け効果係数（冷房） | float | - |
| 289 | nac_shade_coef_heating | 日除け効果係数（暖房） | float | - |
| 290 | nac_wall_name | 外壁名称 | str | 転記 |
| 291 | nac_wall_area | 外皮面積（窓含） | float | ㎡ |
| 292 | nac_window_name | 窓名称 | str | 転記 |
| 293 | nac_window_area | 窓面積 | float | ㎡ |
| 294 | nac_has_blind | ブラインドの有無 | str | 選択 |
| 295 | nac_note | 備考 | str | |

---

## 合計列数

| カテゴリ | 列数 |
|----------|------|
| 共通列 | 10 |
| 様式1（室） | 14 |
| 様式2-1（ゾーン） | 12 |
| 様式2-2（外壁） | 9 |
| 様式2-3（窓） | 8 |
| 様式2-4（外皮） | 11 |
| 様式2-5（熱源） | 24 |
| 様式2-6（ポンプ） | 11 |
| 様式2-7（空調機） | 26 |
| 様式2-8（熱源水温度） | 13 |
| 様式2-9（全熱交換器） | 20 |
| 様式2-10（VWVポンプ） | 10 |
| 様式2-11（PAC特性） | 9 |
| 様式3-1（換気室） | 8 |
| 様式3-2（換気送風機） | 7 |
| 様式3-3（換気空調機） | 12 |
| 様式3-4（負荷率） | 3 |
| 様式4（照明） | 18 |
| 様式5-1（給湯室） | 9 |
| 様式5-2（給湯機器） | 10 |
| 様式6（昇降機） | 11 |
| 様式7-1（太陽光） | 8 |
| 様式7-3（CGS） | 17 |
| 様式8（非空調外皮） | 15 |
| **合計** | **約295列** |

---

## データ例（実際の見え方）

```
file_id | building_name | region | entity_type | room_floor | room_name | room_area | hs_group_name | hs_type | lt_fixture_name | lt_fixture_power | ...（295列）
--------|---------------|--------|-------------|------------|-----------|-----------|---------------|---------|-----------------|------------------|----
001     | 〇〇ビル      | 6      | room        | 1F         | 事務室A   | 120.5     |               |         |                 |                  |
001     | 〇〇ビル      | 6      | room        | 1F         | 会議室    | 30.0      |               |         |                 |                  |
001     | 〇〇ビル      | 6      | room        | 2F         | 事務室B   | 150.0     |               |         |                 |                  |
001     | 〇〇ビル      | 6      | zone        | 1F         |           |           |               |         |                 |                  |
001     | 〇〇ビル      | 6      | heatsource  |            |           |           | 熱源群1       | 空冷チラー |               |                  |
001     | 〇〇ビル      | 6      | heatsource  |            |           |           | 熱源群1       | 空冷チラー |               |                  |
001     | 〇〇ビル      | 6      | ahu         |            |           |           |               |         |                 |                  |
001     | 〇〇ビル      | 6      | lighting    | 1F         | 事務室A   |           |               |         | LED-A           | 40               |
001     | 〇〇ビル      | 6      | lighting    | 1F         | 事務室A   |           |               |         | ダウンライト    | 20               |
001     | 〇〇ビル      | 6      | lighting    | 1F         | 会議室    |           |               |         | LED-B           | 32               |
001     | 〇〇ビル      | 6      | elevator    |            |           |           |               |         |                 |                  |
001     | 〇〇ビル      | 6      | pv          |            |           |           |               |         |                 |                  |
002     | △△ビル      | 6      | room        | 1F         | ロビー    | 80.0      |               |         |                 |                  |
...     | ...           | ...    | ...         | ...        | ...       | ...       | ...           | ...     | ...             | ...              |
```

---

## Python処理例

```python
import pandas as pd

# 読み込み
df = pd.read_excel('webpro_all_data.xlsx')

# 基本情報
print(f"総行数: {len(df)}")
print(f"総列数: {len(df.columns)}")
print(f"建物数: {df['file_id'].nunique()}")

# entity_type別に抽出
df_rooms = df[df['entity_type'] == 'room']
df_hs = df[df['entity_type'] == 'heatsource']
df_ahu = df[df['entity_type'] == 'ahu']
df_light = df[df['entity_type'] == 'lighting']
df_pv = df[df['entity_type'] == 'pv']

# 建物ごとの延床面積
total_area = df_rooms.groupby('file_id')['room_area'].sum()

# 熱源種別の分布
hs_dist = df_hs['hs_type'].value_counts()

# 照明電力密度（LPD）計算
light_summary = df_light.groupby('file_id').agg({
    'lt_fixture_power': lambda x: (x * df_light.loc[x.index, 'lt_fixture_count']).sum()
})
lpd = light_summary['lt_fixture_power'] / total_area

# 特定建物の全データ
building_001 = df[df['file_id'] == '001']
```

---

## まとめ

| 項目 | 内容 |
|------|------|
| 方式 | 縦展開（1レコード=1行） |
| 行数 | 約10,000〜40,000行 |
| 列数 | **295列** |
| キー | file_id + entity_type |
| 特徴 | 全データ保持、該当外列はNULL |
