# -*- coding: utf-8 -*-
import os
import sys
import shutil

# ============ 开发环境配置 ============
# 设置为 True 使用本地开发版本的 PySap2000
DEV_MODE = False

if DEV_MODE:
    # 获取当前脚本所在目录的父目录（即工作区根目录）
    _script_dir = os.path.dirname(os.path.abspath(__file__))
    _workspace_root = os.path.dirname(_script_dir)
    
    # 将工作区根目录添加到 sys.path 最前面，确保优先导入本地开发版本
    if _workspace_root not in sys.path:
        sys.path.insert(0, _workspace_root)
    print(f"[DEV] Using local development version: {_workspace_root}")
# =====================================

import rhinoscriptsyntax as rs
from PySap2000.visualization.rhino import rhino_utils
from PySap2000.global_parameters import Units, UnitSystem


# SAP2000 长度单位到 Rhino 单位系统的映射
SAP_TO_RHINO_UNIT = {
    "mm": 2,   # Rhino: Millimeters
    "cm": 3,   # Rhino: Centimeters
    "m": 4,    # Rhino: Meters
    "in": 8,   # Rhino: Inches
    "ft": 9,   # Rhino: Feet
}

# Rhino 单位代码到 SAP2000 单位系统的映射 (使用 kN 作为力单位)
RHINO_TO_SAP_UNIT = {
    2: UnitSystem.KN_MM_C,   # Millimeters -> kN-mm-C
    3: UnitSystem.KN_CM_C,   # Centimeters -> kN-cm-C
    4: UnitSystem.KN_M_C,    # Meters -> kN-m-C
    8: UnitSystem.KIP_IN_F,  # Inches -> kip-in-F
    9: UnitSystem.KIP_FT_F,  # Feet -> kip-ft-F
}

# Rhino 单位系统代码到名称
RHINO_UNIT_NAMES = {
    0: "None",
    1: "Microns",
    2: "Millimeters",
    3: "Centimeters",
    4: "Meters",
    5: "Kilometers",
    6: "Microinches",
    7: "Mils",
    8: "Inches",
    9: "Feet",
    10: "Miles",
}


def clear_comtypes_cache():
    """清除 comtypes 缓存，解决不同 SAP2000 版本的兼容问题"""
    cache_path = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'comtypes_cache')
    if os.path.exists(cache_path):
        shutil.rmtree(cache_path)


def sync_sap_units_to_rhino(model) -> tuple[bool, str]:
    """
    Sync SAP2000 units to match Rhino units
    
    Returns:
        (success, message): success flag and status message
    """
    # Get Rhino current units
    rhino_unit_code = rs.UnitSystem()
    rhino_unit_name = RHINO_UNIT_NAMES.get(rhino_unit_code, f"Unknown({rhino_unit_code})")
    
    # Get SAP2000 current units
    sap_units = Units.get_present_units(model)
    sap_unit_desc = Units.get_unit_description(sap_units)
    sap_length_unit = Units.get_length_unit(sap_units)
    
    # Check if already matched
    expected_rhino_code = SAP_TO_RHINO_UNIT.get(sap_length_unit)
    if expected_rhino_code == rhino_unit_code:
        return True, f"SAP2000: {sap_unit_desc}  |  Rhino: {rhino_unit_name}"
    
    # Find target SAP2000 unit
    target_sap_unit = RHINO_TO_SAP_UNIT.get(rhino_unit_code)
    if target_sap_unit is None:
        return False, f"Rhino unit [{rhino_unit_name}] not supported"
    
    # Set SAP2000 units
    result = Units.set_present_units(model, target_sap_unit)
    if result != 0:
        return False, f"Failed to set SAP2000 units (code: {result})"
    
    new_unit_desc = Units.get_unit_description(target_sap_unit)
    return True, f"SAP2000: {sap_unit_desc} -> {new_unit_desc}\nRhino: {rhino_unit_name}"


def get_sap_model(retry_on_fail=True):
    try:
        import time
        from PySap2000.application import Application
        
        total_steps = 5
        start_time = time.time()
        
        # [1/5] 连接 SAP2000
        print(f"[1/{total_steps}] Connecting SAP2000...")
        app = Application()
        
        model_path = app.model.GetModelFilepath()
        if not model_path:
            rs.MessageBox("Please save SAP2000 model first", 16, "Info")
            return None
        
        model_filename = app.model.GetModelFilename(False)
        model_name = os.path.splitext(model_filename)[0]
        print(f"  Model: {model_name}")
        
        # [2/5] 同步单位
        print(f"[2/{total_steps}] Syncing units...")
        original_sap_units = Units.get_present_units(app.model)
        success, unit_msg = sync_sap_units_to_rhino(app.model)
        if not success:
            rs.MessageBox(f"Unit sync failed:\n{unit_msg}", 16, "Error")
            return None
        print(f"  {unit_msg}")
        
        # [3/5] 提取模型数据
        print(f"[3/{total_steps}] Extracting model data...")
        json_file = os.path.join(model_path, f"{model_name}_model_data.json")
        result = rhino_utils.export_to_json(app.model, json_file, unit_scale=1.0)
        
        # [4/5] 恢复单位
        print(f"[4/{total_steps}] Restoring SAP2000 units...")
        Units.set_present_units(app.model, original_sap_units)
        
        # [5/5] 完成
        elapsed = time.time() - start_time
        print(f"[5/{total_steps}] Done! ({elapsed:.1f}s)")
        
        rs.MessageBox(
            f"Model [{model_name}] extracted!\n\n"
            f"{unit_msg}\n"
            f"Time: {elapsed:.1f}s",
            0, "Success"
        )
        
        return result
        
    except Exception as e:
        error_msg = str(e)
        
        # Clear comtypes cache and retry on COM interface error
        if retry_on_fail and ("不支持此接口" in error_msg or "-2147467262" in error_msg):
            clear_comtypes_cache()
            return get_sap_model(retry_on_fail=False)
        
        rs.MessageBox(f"Error: {error_msg}", 16, "Error")
        return None


if __name__ == "__main__":
    get_sap_model()
    
