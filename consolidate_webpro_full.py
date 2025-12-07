#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WEBPRO入力シート統合スクリプト
100ファイルのWEBPRO入力シートを1シート（295列）に統合

使用方法:
    python consolidate_webpro_full.py --input_dir ./webpro_files --output ./webpro_all_data.xlsx
"""

import pandas as pd
import numpy as np
from pathlib import Path
import argparse
from typing import Dict, List, Any, Optional
import warnings
warnings.filterwarnings('ignore')

# =============================================================================
# 列定義
# =============================================================================

# 共通列（全行に付与）
COMMON_COLUMNS = [
    'file_id',              # ファイル識別子
    'building_name',        # 建物の名称
    'prefecture',           # 都道府県
    'city',                 # 市区町村
    'region',               # 地域区分
    'structure',            # 構造
    'floors_above',         # 地上階数
    'floors_below',         # 地下階数
    'evaluation_target',    # 評価対象
    'entity_type',          # データ種別
]

# 様式1: 室仕様
ROOM_COLUMNS = [
    'room_floor', 'room_name', 'room_building_type',
    'room_type_major', 'room_type_minor', 'room_area',
    'room_floor_height', 'room_ceiling_height',
    'room_is_ac_target', 'room_is_vent_target',
    'room_is_light_target', 'room_is_hotwater_target',
    'room_building_group', 'room_note',
]

# 様式2-1: 空調ゾーン
ZONE_COLUMNS = [
    'zone_floor', 'zone_room_name', 'zone_room_type_major', 'zone_room_type_minor',
    'zone_room_area', 'zone_floor_height', 'zone_ceiling_height',
    'zone_ac_floor', 'zone_name', 'zone_ahu_group_room', 'zone_ahu_group_oa', 'zone_note',
]

# 様式2-2: 外壁構成
WALL_COLUMNS = [
    'wall_name', 'wall_type', 'wall_u_value', 'wall_material_no',
    'wall_material_name', 'wall_conductivity', 'wall_thickness',
    'wall_solar_absorption', 'wall_note',
]

# 様式2-3: 窓仕様
WINDOW_COLUMNS = [
    'window_name', 'window_u_value', 'window_eta_value',
    'window_frame_type', 'window_glass_type',
    'window_glass_u_value', 'window_glass_eta_value', 'window_note',
]

# 様式2-4: 外皮
ENVELOPE_COLUMNS = [
    'env_floor', 'env_zone_name', 'env_direction',
    'env_shade_coef_cooling', 'env_shade_coef_heating',
    'env_wall_name', 'env_wall_area', 'env_window_name',
    'env_window_area', 'env_has_blind', 'env_note',
]

# 様式2-5: 熱源
HEATSOURCE_COLUMNS = [
    'hs_group_name', 'hs_simultaneous', 'hs_staging_control',
    'hs_operation_mode', 'hs_storage_capacity', 'hs_type',
    'hs_cooling_order', 'hs_cooling_count', 'hs_cooling_supply_temp',
    'hs_cooling_capacity', 'hs_cooling_main_power', 'hs_cooling_sub_power',
    'hs_cooling_pump_power', 'hs_ct_capacity', 'hs_ct_fan_power', 'hs_ct_pump_power',
    'hs_heating_order', 'hs_heating_count', 'hs_heating_supply_temp',
    'hs_heating_capacity', 'hs_heating_main_power', 'hs_heating_sub_power',
    'hs_heating_pump_power', 'hs_note',
]

# 様式2-6: 二次ポンプ
PUMP_COLUMNS = [
    'pump_group_name', 'pump_staging_control', 'pump_cooling_temp_diff',
    'pump_heating_temp_diff', 'pump_order', 'pump_count',
    'pump_rated_flow', 'pump_rated_power', 'pump_flow_control',
    'pump_min_flow_ratio', 'pump_note',
]

# 様式2-7: 空調機
AHU_COLUMNS = [
    'ahu_group_name', 'ahu_count', 'ahu_type',
    'ahu_cooling_capacity', 'ahu_heating_capacity', 'ahu_oa_flow',
    'ahu_sa_fan_power', 'ahu_ra_fan_power', 'ahu_oa_fan_power', 'ahu_ea_fan_power',
    'ahu_air_flow_control', 'ahu_min_air_ratio',
    'ahu_preheat_oa_stop', 'ahu_economizer', 'ahu_has_hex', 'ahu_hex_name',
    'ahu_hex_flow', 'ahu_hex_eff_cooling', 'ahu_hex_eff_heating',
    'ahu_auto_bypass', 'ahu_rotor_power',
    'ahu_pump_group_cooling', 'ahu_pump_group_heating',
    'ahu_hs_group_cooling', 'ahu_hs_group_heating', 'ahu_note',
]

# 様式2-8: 熱源水温度
HS_WATER_TEMP_COLUMNS = [
    'hswt_group_name',
    'hswt_temp_jan', 'hswt_temp_feb', 'hswt_temp_mar', 'hswt_temp_apr',
    'hswt_temp_may', 'hswt_temp_jun', 'hswt_temp_jul', 'hswt_temp_aug',
    'hswt_temp_sep', 'hswt_temp_oct', 'hswt_temp_nov', 'hswt_temp_dec',
]

# 様式2-9: 全熱交換器
HEAT_EXCHANGER_COLUMNS = [
    'hex_name', 'hex_type', 'hex_oa_flow', 'hex_ea_flow', 'hex_count',
    'hex_eff_cooling_1', 'hex_eff_heating_1', 'hex_test_sa_flow_1',
    'hex_test_ra_flow_1', 'hex_vent_eff_1',
    'hex_eff_cooling_2', 'hex_eff_heating_2', 'hex_test_sa_flow_2',
    'hex_test_ra_flow_2', 'hex_vent_eff_2',
    'hex_eff_cooling_3', 'hex_eff_heating_3', 'hex_test_sa_flow_3',
    'hex_test_ra_flow_3', 'hex_vent_eff_3',
]

# 様式2-10: 変流量二次ポンプ
VWV_PUMP_COLUMNS = [
    'vwv_group_name', 'vwv_cooling_temp_diff', 'vwv_heating_temp_diff',
    'vwv_rated_flow', 'vwv_rated_power', 'vwv_min_flow_ratio',
    'vwv_coef_3rd', 'vwv_coef_2nd', 'vwv_coef_1st', 'vwv_coef_const',
]

# 様式2-11: PAC部分負荷特性
PAC_PARTIAL_COLUMNS = [
    'pac_hs_name',
    'pac_cooling_coef_2nd', 'pac_cooling_coef_1st', 'pac_cooling_const', 'pac_cooling_min_output',
    'pac_heating_coef_2nd', 'pac_heating_coef_1st', 'pac_heating_const', 'pac_heating_min_output',
]

# 様式3-1: 換気室
VENT_ROOM_COLUMNS = [
    'vr_floor', 'vr_room_name', 'vr_room_type_major', 'vr_room_type_minor',
    'vr_room_area', 'vr_vent_type', 'vr_vent_equip_name', 'vr_note',
]

# 様式3-2: 換気送風機
VENT_FAN_COLUMNS = [
    'vf_equip_name', 'vf_design_flow', 'vf_motor_power',
    'vf_high_eff_motor', 'vf_has_inverter', 'vf_flow_control', 'vf_note',
]

# 様式3-3: 換気空調機
VENT_AHU_COLUMNS = [
    'va_equip_name', 'va_room_type', 'va_cooling_capacity',
    'va_hs_efficiency', 'va_pump_power', 'va_fan_type',
    'va_design_flow', 'va_motor_power', 'va_high_eff_motor',
    'va_has_inverter', 'va_flow_control', 'va_note',
]

# 様式3-4: 年間平均負荷率
VENT_LOAD_RATE_COLUMNS = [
    'vlr_equip_name', 'vlr_annual_load_rate', 'vlr_note',
]

# 様式4: 照明
LIGHTING_COLUMNS = [
    'lt_floor', 'lt_room_name', 'lt_room_type_major', 'lt_room_type_minor',
    'lt_room_area', 'lt_floor_height', 'lt_ceiling_height',
    'lt_room_width', 'lt_room_depth', 'lt_room_index',
    'lt_fixture_name', 'lt_fixture_power', 'lt_fixture_count',
    'lt_occupancy_control', 'lt_daylight_control',
    'lt_schedule_control', 'lt_initial_correction', 'lt_note',
]

# 様式5-1: 給湯室
HOTWATER_ROOM_COLUMNS = [
    'hwr_floor', 'hwr_room_name', 'hwr_room_type_major', 'hwr_room_type_minor',
    'hwr_room_area', 'hwr_supply_location', 'hwr_water_saving',
    'hwr_equip_name', 'hwr_note',
]

# 様式5-2: 給湯機器
HOTWATER_EQUIP_COLUMNS = [
    'hwe_equip_name', 'hwe_fuel_type', 'hwe_heating_capacity',
    'hwe_efficiency', 'hwe_insulation', 'hwe_pipe_diameter',
    'hwe_solar_area', 'hwe_solar_azimuth', 'hwe_solar_tilt', 'hwe_note',
]

# 様式6: 昇降機
ELEVATOR_COLUMNS = [
    'ev_floor', 'ev_room_name', 'ev_room_type_major', 'ev_room_type_minor',
    'ev_equip_name', 'ev_count', 'ev_capacity', 'ev_speed',
    'ev_transport_coef', 'ev_control_type', 'ev_note',
]

# 様式7-1: 太陽光発電
PV_COLUMNS = [
    'pv_system_name', 'pv_pcs_efficiency', 'pv_cell_type',
    'pv_install_type', 'pv_capacity', 'pv_azimuth', 'pv_tilt', 'pv_note',
]

# 様式7-3: コージェネレーション
CGS_COLUMNS = [
    'cgs_name', 'cgs_rated_output', 'cgs_count',
    'cgs_gen_eff_100', 'cgs_gen_eff_75', 'cgs_gen_eff_50',
    'cgs_heat_eff_100', 'cgs_heat_eff_75', 'cgs_heat_eff_50',
    'cgs_priority_ac_cool', 'cgs_priority_ac_heat', 'cgs_priority_hotwater',
    'cgs_24h_operation', 'cgs_ac_cool_hs_group', 'cgs_ac_heat_hs_group',
    'cgs_hotwater_equip', 'cgs_note',
]

# 様式8: 非空調外皮
ENVELOPE_NON_AC_COLUMNS = [
    'nac_floor', 'nac_zone_name', 'nac_room_type_major', 'nac_room_type_minor',
    'nac_room_area', 'nac_floor_height', 'nac_direction',
    'nac_shade_coef_cooling', 'nac_shade_coef_heating',
    'nac_wall_name', 'nac_wall_area', 'nac_window_name',
    'nac_window_area', 'nac_has_blind', 'nac_note',
]

# 全列リスト
ALL_COLUMNS = (
    COMMON_COLUMNS +
    ROOM_COLUMNS +
    ZONE_COLUMNS +
    WALL_COLUMNS +
    WINDOW_COLUMNS +
    ENVELOPE_COLUMNS +
    HEATSOURCE_COLUMNS +
    PUMP_COLUMNS +
    AHU_COLUMNS +
    HS_WATER_TEMP_COLUMNS +
    HEAT_EXCHANGER_COLUMNS +
    VWV_PUMP_COLUMNS +
    PAC_PARTIAL_COLUMNS +
    VENT_ROOM_COLUMNS +
    VENT_FAN_COLUMNS +
    VENT_AHU_COLUMNS +
    VENT_LOAD_RATE_COLUMNS +
    LIGHTING_COLUMNS +
    HOTWATER_ROOM_COLUMNS +
    HOTWATER_EQUIP_COLUMNS +
    ELEVATOR_COLUMNS +
    PV_COLUMNS +
    CGS_COLUMNS +
    ENVELOPE_NON_AC_COLUMNS
)

# =============================================================================
# 様式設定（シート名、データ開始行、列マッピング）
# =============================================================================

SHEET_CONFIG = {
    'room': {
        'sheet_name': '1) 室仕様',
        'data_start_row': 9,
        'columns': ROOM_COLUMNS,
        'col_mapping': {
            0: 'room_floor',
            1: 'room_name',
            2: 'room_building_type',
            3: 'room_type_major',
            4: 'room_type_minor',
            5: 'room_area',
            6: 'room_floor_height',
            7: 'room_ceiling_height',
            8: 'room_is_ac_target',
            9: 'room_is_vent_target',
            10: 'room_is_light_target',
            11: 'room_is_hotwater_target',
            12: 'room_building_group',
            13: 'room_note',
        }
    },
    'zone': {
        'sheet_name': '2-1) 空調ゾーン',
        'data_start_row': 9,
        'columns': ZONE_COLUMNS,
        'col_mapping': {
            0: 'zone_floor',
            1: 'zone_room_name',
            2: 'zone_room_type_major',
            3: 'zone_room_type_minor',
            4: 'zone_room_area',
            5: 'zone_floor_height',
            6: 'zone_ceiling_height',
            7: 'zone_ac_floor',
            8: 'zone_name',
            9: 'zone_ahu_group_room',
            10: 'zone_ahu_group_oa',
            # 11-13: 室用途選択リスト（スキップ）
            14: 'zone_note',
        }
    },
    'wall': {
        'sheet_name': '2-2) 外壁構成 ',
        'data_start_row': 9,
        'columns': WALL_COLUMNS,
        'col_mapping': {
            0: 'wall_name',
            1: 'wall_type',
            2: 'wall_u_value',
            3: 'wall_material_no',
            4: 'wall_material_name',
            5: 'wall_conductivity',
            6: 'wall_thickness',
            7: 'wall_solar_absorption',
            8: 'wall_note',
        }
    },
    'window': {
        'sheet_name': '2-3) 窓仕様',
        'data_start_row': 9,
        'columns': WINDOW_COLUMNS,
        'col_mapping': {
            0: 'window_name',
            1: 'window_u_value',
            2: 'window_eta_value',
            3: 'window_frame_type',
            4: 'window_glass_type',
            5: 'window_glass_u_value',
            6: 'window_glass_eta_value',
            7: 'window_note',
        }
    },
    'envelope': {
        'sheet_name': '2-4) 外皮 ',
        'data_start_row': 9,
        'columns': ENVELOPE_COLUMNS,
        'col_mapping': {
            0: 'env_floor',
            1: 'env_zone_name',
            2: 'env_direction',
            3: 'env_shade_coef_cooling',
            4: 'env_shade_coef_heating',
            5: 'env_wall_name',
            6: 'env_wall_area',
            7: 'env_window_name',
            8: 'env_window_area',
            9: 'env_has_blind',
            10: 'env_note',
        }
    },
    'heatsource': {
        'sheet_name': '2-5) 熱源',
        'data_start_row': 9,
        'columns': HEATSOURCE_COLUMNS,
        'col_mapping': {
            0: 'hs_group_name',
            1: 'hs_simultaneous',
            2: 'hs_staging_control',
            3: 'hs_operation_mode',
            4: 'hs_storage_capacity',
            5: 'hs_type',
            6: 'hs_cooling_order',
            7: 'hs_cooling_count',
            8: 'hs_cooling_supply_temp',
            9: 'hs_cooling_capacity',
            10: 'hs_cooling_main_power',
            11: 'hs_cooling_sub_power',
            12: 'hs_cooling_pump_power',
            13: 'hs_ct_capacity',
            14: 'hs_ct_fan_power',
            15: 'hs_ct_pump_power',
            16: 'hs_heating_order',
            17: 'hs_heating_count',
            18: 'hs_heating_supply_temp',
            19: 'hs_heating_capacity',
            20: 'hs_heating_main_power',
            21: 'hs_heating_sub_power',
            22: 'hs_heating_pump_power',
            23: 'hs_note',
        }
    },
    'pump': {
        'sheet_name': '2-6) 2次ﾎﾟﾝﾌﾟ',
        'data_start_row': 9,
        'columns': PUMP_COLUMNS,
        'col_mapping': {
            0: 'pump_group_name',
            1: 'pump_staging_control',
            2: 'pump_cooling_temp_diff',
            3: 'pump_heating_temp_diff',
            4: 'pump_order',
            5: 'pump_count',
            6: 'pump_rated_flow',
            7: 'pump_rated_power',
            8: 'pump_flow_control',
            9: 'pump_min_flow_ratio',
            10: 'pump_note',
        }
    },
    'ahu': {
        'sheet_name': '2-7) 空調機',
        'data_start_row': 9,
        'columns': AHU_COLUMNS,
        'col_mapping': {
            0: 'ahu_group_name',
            1: 'ahu_count',
            2: 'ahu_type',
            3: 'ahu_cooling_capacity',
            4: 'ahu_heating_capacity',
            5: 'ahu_oa_flow',
            6: 'ahu_sa_fan_power',
            7: 'ahu_ra_fan_power',
            8: 'ahu_oa_fan_power',
            9: 'ahu_ea_fan_power',
            10: 'ahu_air_flow_control',
            11: 'ahu_min_air_ratio',
            12: 'ahu_preheat_oa_stop',
            13: 'ahu_economizer',
            14: 'ahu_has_hex',
            15: 'ahu_hex_name',
            16: 'ahu_hex_flow',
            17: 'ahu_hex_eff_cooling',
            18: 'ahu_hex_eff_heating',
            19: 'ahu_auto_bypass',
            20: 'ahu_rotor_power',
            21: 'ahu_pump_group_cooling',
            22: 'ahu_pump_group_heating',
            23: 'ahu_hs_group_cooling',
            24: 'ahu_hs_group_heating',
            25: 'ahu_note',
        }
    },
    'hs_water_temp': {
        'sheet_name': '2-8) 熱源水温度',
        'data_start_row': 9,
        'columns': HS_WATER_TEMP_COLUMNS,
        'col_mapping': {
            0: 'hswt_group_name',
            1: 'hswt_temp_jan',
            2: 'hswt_temp_feb',
            3: 'hswt_temp_mar',
            4: 'hswt_temp_apr',
            5: 'hswt_temp_may',
            6: 'hswt_temp_jun',
            7: 'hswt_temp_jul',
            8: 'hswt_temp_aug',
            9: 'hswt_temp_sep',
            10: 'hswt_temp_oct',
            11: 'hswt_temp_nov',
            12: 'hswt_temp_dec',
        }
    },
    'heat_exchanger': {
        'sheet_name': '2-9) 全熱交換器',
        'data_start_row': 9,
        'columns': HEAT_EXCHANGER_COLUMNS,
        'col_mapping': {
            0: 'hex_name',
            1: 'hex_type',
            2: 'hex_oa_flow',
            3: 'hex_ea_flow',
            4: 'hex_count',
            5: 'hex_eff_cooling_1',
            6: 'hex_eff_heating_1',
            7: 'hex_test_sa_flow_1',
            8: 'hex_test_ra_flow_1',
            9: 'hex_vent_eff_1',
            10: 'hex_eff_cooling_2',
            11: 'hex_eff_heating_2',
            12: 'hex_test_sa_flow_2',
            13: 'hex_test_ra_flow_2',
            14: 'hex_vent_eff_2',
            15: 'hex_eff_cooling_3',
            16: 'hex_eff_heating_3',
            17: 'hex_test_sa_flow_3',
            18: 'hex_test_ra_flow_3',
            19: 'hex_vent_eff_3',
        }
    },
    'vwv_pump': {
        'sheet_name': '2-10) 変流量二次ポンプシステム',
        'data_start_row': 8,
        'columns': VWV_PUMP_COLUMNS,
        'col_mapping': {
            0: 'vwv_group_name',
            1: 'vwv_cooling_temp_diff',
            2: 'vwv_heating_temp_diff',
            3: 'vwv_rated_flow',
            4: 'vwv_rated_power',
            5: 'vwv_min_flow_ratio',
            6: 'vwv_coef_3rd',
            7: 'vwv_coef_2nd',
            8: 'vwv_coef_1st',
            9: 'vwv_coef_const',
        }
    },
    'pac_partial': {
        'sheet_name': '2-11) パッケージエアコンディショナ(空冷式)部分負荷特性',
        'data_start_row': 5,
        'columns': PAC_PARTIAL_COLUMNS,
        'col_mapping': {
            0: 'pac_hs_name',
            1: 'pac_cooling_coef_2nd',
            2: 'pac_cooling_coef_1st',
            3: 'pac_cooling_const',
            4: 'pac_cooling_min_output',
            5: 'pac_heating_coef_2nd',
            6: 'pac_heating_coef_1st',
            7: 'pac_heating_const',
            8: 'pac_heating_min_output',
        }
    },
    'vent_room': {
        'sheet_name': '3-1) 換気室',
        'data_start_row': 9,
        'columns': VENT_ROOM_COLUMNS,
        'col_mapping': {
            0: 'vr_floor',
            1: 'vr_room_name',
            2: 'vr_room_type_major',
            3: 'vr_room_type_minor',
            4: 'vr_room_area',
            5: 'vr_vent_type',
            6: 'vr_vent_equip_name',
            7: 'vr_note',
        }
    },
    'vent_fan': {
        'sheet_name': '3-2) 換気送風機',
        'data_start_row': 9,
        'columns': VENT_FAN_COLUMNS,
        'col_mapping': {
            0: 'vf_equip_name',
            1: 'vf_design_flow',
            2: 'vf_motor_power',
            3: 'vf_high_eff_motor',
            4: 'vf_has_inverter',
            5: 'vf_flow_control',
            6: 'vf_note',
        }
    },
    'vent_ahu': {
        'sheet_name': '3-3) 換気空調機',
        'data_start_row': 8,
        'columns': VENT_AHU_COLUMNS,
        'col_mapping': {
            0: 'va_equip_name',
            1: 'va_room_type',
            2: 'va_cooling_capacity',
            3: 'va_hs_efficiency',
            4: 'va_pump_power',
            5: 'va_fan_type',
            6: 'va_design_flow',
            7: 'va_motor_power',
            8: 'va_high_eff_motor',
            9: 'va_has_inverter',
            10: 'va_flow_control',
            11: 'va_note',
        }
    },
    'vent_load_rate': {
        'sheet_name': '3-4) 年間平均負荷率',
        'data_start_row': 9,
        'columns': VENT_LOAD_RATE_COLUMNS,
        'col_mapping': {
            0: 'vlr_equip_name',
            1: 'vlr_annual_load_rate',
            2: 'vlr_note',
        }
    },
    'lighting': {
        'sheet_name': '4) 照明',
        'data_start_row': 9,
        'columns': LIGHTING_COLUMNS,
        'col_mapping': {
            0: 'lt_floor',
            1: 'lt_room_name',
            2: 'lt_room_type_major',
            3: 'lt_room_type_minor',
            4: 'lt_room_area',
            5: 'lt_floor_height',
            6: 'lt_ceiling_height',
            7: 'lt_room_width',
            8: 'lt_room_depth',
            9: 'lt_room_index',
            10: 'lt_fixture_name',
            11: 'lt_fixture_power',
            12: 'lt_fixture_count',
            13: 'lt_occupancy_control',
            14: 'lt_daylight_control',
            15: 'lt_schedule_control',
            16: 'lt_initial_correction',
            17: 'lt_note',
        }
    },
    'hotwater_room': {
        'sheet_name': '5-1) 給湯室',
        'data_start_row': 9,
        'columns': HOTWATER_ROOM_COLUMNS,
        'col_mapping': {
            0: 'hwr_floor',
            1: 'hwr_room_name',
            2: 'hwr_room_type_major',
            3: 'hwr_room_type_minor',
            4: 'hwr_room_area',
            5: 'hwr_supply_location',
            6: 'hwr_water_saving',
            7: 'hwr_equip_name',
            8: 'hwr_note',
        }
    },
    'hotwater_equip': {
        'sheet_name': '5-2) 給湯機器',
        'data_start_row': 9,
        'columns': HOTWATER_EQUIP_COLUMNS,
        'col_mapping': {
            0: 'hwe_equip_name',
            1: 'hwe_fuel_type',
            2: 'hwe_heating_capacity',
            3: 'hwe_efficiency',
            4: 'hwe_insulation',
            5: 'hwe_pipe_diameter',
            6: 'hwe_solar_area',
            7: 'hwe_solar_azimuth',
            8: 'hwe_solar_tilt',
            9: 'hwe_note',
        }
    },
    'elevator': {
        'sheet_name': '6) 昇降機',
        'data_start_row': 9,
        'columns': ELEVATOR_COLUMNS,
        'col_mapping': {
            0: 'ev_floor',
            1: 'ev_room_name',
            2: 'ev_room_type_major',
            3: 'ev_room_type_minor',
            4: 'ev_equip_name',
            5: 'ev_count',
            6: 'ev_capacity',
            7: 'ev_speed',
            8: 'ev_transport_coef',
            9: 'ev_control_type',
            10: 'ev_note',
        }
    },
    'pv': {
        'sheet_name': '7-1) 太陽光発電',
        'data_start_row': 9,
        'columns': PV_COLUMNS,
        'col_mapping': {
            0: 'pv_system_name',
            1: 'pv_pcs_efficiency',
            2: 'pv_cell_type',
            3: 'pv_install_type',
            4: 'pv_capacity',
            5: 'pv_azimuth',
            6: 'pv_tilt',
            7: 'pv_note',
        }
    },
    'cgs': {
        'sheet_name': '7-3) コージェネレーション設備',
        'data_start_row': 8,
        'columns': CGS_COLUMNS,
        'col_mapping': {
            0: 'cgs_name',
            1: 'cgs_rated_output',
            2: 'cgs_count',
            3: 'cgs_gen_eff_100',
            4: 'cgs_gen_eff_75',
            5: 'cgs_gen_eff_50',
            6: 'cgs_heat_eff_100',
            7: 'cgs_heat_eff_75',
            8: 'cgs_heat_eff_50',
            9: 'cgs_priority_ac_cool',
            10: 'cgs_priority_ac_heat',
            11: 'cgs_priority_hotwater',
            12: 'cgs_24h_operation',
            13: 'cgs_ac_cool_hs_group',
            14: 'cgs_ac_heat_hs_group',
            15: 'cgs_hotwater_equip',
            16: 'cgs_note',
        }
    },
    'envelope_non_ac': {
        'sheet_name': '8) 非空調外皮',
        'data_start_row': 9,
        'columns': ENVELOPE_NON_AC_COLUMNS,
        'col_mapping': {
            0: 'nac_floor',
            1: 'nac_zone_name',
            2: 'nac_room_type_major',
            3: 'nac_room_type_minor',
            4: 'nac_room_area',
            5: 'nac_floor_height',
            6: 'nac_direction',
            7: 'nac_shade_coef_cooling',
            8: 'nac_shade_coef_heating',
            9: 'nac_wall_name',
            10: 'nac_wall_area',
            11: 'nac_window_name',
            12: 'nac_window_area',
            13: 'nac_has_blind',
            14: 'nac_note',
        }
    },
}


# =============================================================================
# 基本情報抽出（様式0）
# =============================================================================

def extract_basic_info(xlsx_path: str) -> Dict[str, Any]:
    """
    様式0から基本情報を抽出
    
    様式0の構造（Rev.2）:
    - Row7: ③評価対象 → Col2に値
    - Row9: ④建物の名称 → Col2に値
    - Row10: ⑤建築物所在地 → Col3:都道府県, Col4以降:市区町村
    - Row12: ⑥省エネ基準地域区分 → Col2に値
    - Row13: ⑦構造 → Col2に値
    - Row14: ⑧階数 → Col3:地上, Col4以降:地下
    """
    try:
        df = pd.read_excel(xlsx_path, sheet_name='0) 基本情報', header=None)
    except Exception as e:
        print(f"Warning: 基本情報シートの読み込み失敗: {e}")
        return {}
    
    basic_info = {}
    
    def get_val(row, col):
        """安全に値を取得"""
        try:
            if row < df.shape[0] and col < df.shape[1]:
                val = df.iloc[row, col]
                if pd.notna(val) and str(val).strip() != '':
                    return val
        except:
            pass
        return None
    
    # 行ごとに解析（様式0は縦型フォーム）
    for row_idx in range(df.shape[0]):
        label = str(df.iloc[row_idx, 1]) if row_idx < df.shape[0] and 1 < df.shape[1] and pd.notna(df.iloc[row_idx, 1]) else ''
        
        if '評価対象' in label:
            basic_info['evaluation_target'] = get_val(row_idx, 2) or ''
            
        elif '建物の名称' in label:
            basic_info['building_name'] = get_val(row_idx, 2) or ''
            
        elif '建築物所在地' in label or '所在地' in label:
            # 都道府県と市区町村を取得
            # Col2に「都道府県」ラベル、Col3に値、Col4に「市区町村」ラベル、Col5に値
            # または Col3に都道府県値、Col4以降に市区町村
            pref = get_val(row_idx, 3)
            city = get_val(row_idx, 5) or get_val(row_idx, 4)
            basic_info['prefecture'] = pref or ''
            basic_info['city'] = city or ''
            
        elif '地域の区分' in label or '地域区分' in label:
            val = get_val(row_idx, 2)
            if val is not None:
                try:
                    basic_info['region'] = int(float(val))
                except:
                    basic_info['region'] = val
            else:
                basic_info['region'] = ''
                
        elif '構造' in label and '外壁' not in label:
            basic_info['structure'] = get_val(row_idx, 2) or ''
            
        elif '階数' in label:
            # Col2に「地上」ラベル、Col3に値、Col4に「地下」ラベル、Col5に値
            floors_above = get_val(row_idx, 3)
            floors_below = get_val(row_idx, 5) or get_val(row_idx, 4)
            if floors_above is not None:
                try:
                    basic_info['floors_above'] = int(float(floors_above))
                except:
                    basic_info['floors_above'] = floors_above
            else:
                basic_info['floors_above'] = ''
            if floors_below is not None:
                try:
                    basic_info['floors_below'] = int(float(floors_below))
                except:
                    basic_info['floors_below'] = floors_below
            else:
                basic_info['floors_below'] = ''
    
    return basic_info


# =============================================================================
# 様式データ抽出
# =============================================================================

def extract_sheet_data(
    xlsx_path: str,
    entity_type: str,
    config: Dict[str, Any]
) -> List[Dict[str, Any]]:
    """
    指定様式からデータを抽出
    """
    try:
        df = pd.read_excel(xlsx_path, sheet_name=config['sheet_name'], header=None)
    except Exception as e:
        # シートが存在しない場合は空リストを返す
        return []
    
    records = []
    data_start_row = config['data_start_row']
    col_mapping = config['col_mapping']
    
    # データ行を走査
    for row_idx in range(data_start_row, df.shape[0]):
        row_data = {}
        has_data = False
        
        for col_idx, col_name in col_mapping.items():
            if col_idx < df.shape[1]:
                val = df.iloc[row_idx, col_idx]
                if pd.notna(val) and str(val).strip() != '':
                    row_data[col_name] = val
                    has_data = True
                else:
                    row_data[col_name] = None
            else:
                row_data[col_name] = None
        
        # データがある行のみ追加
        if has_data:
            row_data['entity_type'] = entity_type
            records.append(row_data)
    
    return records


# =============================================================================
# 1ファイル処理
# =============================================================================

def process_single_file(xlsx_path: str, file_id: str) -> List[Dict[str, Any]]:
    """
    1つのWEBPROファイルを処理し、全レコードを返す
    """
    all_records = []
    
    # 基本情報を抽出
    basic_info = extract_basic_info(xlsx_path)
    
    # 各様式からデータを抽出
    for entity_type, config in SHEET_CONFIG.items():
        records = extract_sheet_data(xlsx_path, entity_type, config)
        
        for record in records:
            # 共通情報を付与
            record['file_id'] = file_id
            record['building_name'] = basic_info.get('building_name', '')
            record['prefecture'] = basic_info.get('prefecture', '')
            record['city'] = basic_info.get('city', '')
            record['region'] = basic_info.get('region', '')
            record['structure'] = basic_info.get('structure', '')
            record['floors_above'] = basic_info.get('floors_above', '')
            record['floors_below'] = basic_info.get('floors_below', '')
            record['evaluation_target'] = basic_info.get('evaluation_target', '')
            
            all_records.append(record)
    
    return all_records


# =============================================================================
# 全ファイル統合
# =============================================================================

def consolidate_files(
    input_dir: str,
    output_path: str,
    file_pattern: str = '*.xlsx'
) -> pd.DataFrame:
    """
    指定ディレクトリ内の全WEBPROファイルを統合
    """
    input_path = Path(input_dir)
    xlsx_files = sorted(input_path.glob(file_pattern))
    
    if not xlsx_files:
        raise FileNotFoundError(f"No Excel files found in {input_dir}")
    
    print(f"Found {len(xlsx_files)} files to process")
    
    all_records = []
    
    for idx, xlsx_file in enumerate(xlsx_files, start=1):
        file_id = f"{idx:03d}"
        print(f"Processing [{file_id}] {xlsx_file.name}...")
        
        try:
            records = process_single_file(str(xlsx_file), file_id)
            all_records.extend(records)
            print(f"  -> {len(records)} records extracted")
        except Exception as e:
            print(f"  -> Error: {e}")
    
    # DataFrameに変換
    df = pd.DataFrame(all_records)
    
    # 列順序を整理（定義順に並べる）
    existing_columns = [col for col in ALL_COLUMNS if col in df.columns]
    df = df[existing_columns]
    
    # 不足列を追加（NaN）
    for col in ALL_COLUMNS:
        if col not in df.columns:
            df[col] = None
    
    # 最終的な列順序
    df = df[ALL_COLUMNS]
    
    # Excel出力
    print(f"\nWriting to {output_path}...")
    df.to_excel(output_path, index=False, sheet_name='all_data')
    
    print(f"\nDone!")
    print(f"  Total records: {len(df)}")
    print(f"  Total columns: {len(df.columns)}")
    print(f"  Buildings: {df['file_id'].nunique()}")
    
    # entity_type別の集計
    print("\nRecords by entity_type:")
    print(df['entity_type'].value_counts().to_string())
    
    return df


# =============================================================================
# メイン
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description='WEBPRO入力シート統合スクリプト'
    )
    parser.add_argument(
        '--input_dir', '-i',
        required=True,
        help='WEBPROファイルが格納されたディレクトリ'
    )
    parser.add_argument(
        '--output', '-o',
        default='webpro_all_data.xlsx',
        help='出力Excelファイルパス（デフォルト: webpro_all_data.xlsx）'
    )
    parser.add_argument(
        '--pattern', '-p',
        default='*.xlsx',
        help='ファイルパターン（デフォルト: *.xlsx）'
    )
    
    args = parser.parse_args()
    
    consolidate_files(
        input_dir=args.input_dir,
        output_path=args.output,
        file_pattern=args.pattern
    )


if __name__ == '__main__':
    main()
