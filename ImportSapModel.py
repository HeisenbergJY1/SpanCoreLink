# -*- coding: utf-8 -*-
import os
import sys

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


def import_sap_model():
    try:
        from PySap2000.application import Application
        import os
        
        print("Connecting SAP2000...")
        app = Application()
        
        model_path = app.model.GetModelFilepath()
        if not model_path:
            rs.MessageBox("Please save model first", 16, "Info")
            return None
        
        model_filename = app.model.GetModelFilename(False)
        model_name = os.path.splitext(model_filename)[0]
        json_file = os.path.join(model_path, f"{model_name}_model_data.json")
        
        if not os.path.exists(json_file):
            rs.MessageBox("Model data not found. Run GetSapModel first", 16, "Error")
            return None
        
    except Exception as e:
        rs.MessageBox(f"Error: {str(e)}", 16, "Error")
        return None
    
    mode = rs.GetString("Mode", "Solid", ["Solid", "Wireframe"])
    if not mode:
        return None
    
    create_solid = (mode == "Solid")
    
    try:
        layer_name = "SAP2000_Model"
        if rs.IsLayer(layer_name):
            print(f"Deleting old layer: {layer_name}")
            old_objects = rs.ObjectsByLayer(layer_name, select=False)
            if old_objects:
                rs.DeleteObjects(old_objects)
            rs.DeleteLayer(layer_name)
        
        print(f"Importing {mode} model...")
        
        guids = rhino_utils.import_from_json(
            json_file,
            layer_name=layer_name,
            create_solid=create_solid,
            organize_by_group=False,  # 不按组分类
            organize_by_type=True     # 按类型分类（框架单元/索单元）
        )
        
        print(f"Done: {len(guids)} objects")
        rs.MessageBox(f"Imported {len(guids)} objects\nMode: {mode}", 0, "Success")
        
        return guids
        
    except Exception as e:
        rs.MessageBox(f"Error: {str(e)}", 16, "Error")
        return None


if __name__ == "__main__":
    import_sap_model()
