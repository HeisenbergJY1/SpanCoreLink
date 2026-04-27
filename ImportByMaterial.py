# -*- coding: utf-8 -*-
"""
ImportByMaterial.py - 按材料导入 SAP2000 模型到 Rhino

按材料分图层，每个材料图层分配随机颜色
"""
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
import random


def import_by_material():
    """按材料导入 SAP2000 模型"""
    try:
        from PySap2000.application import Application
        
        print("Connecting SAP2000...")
        app = Application()
        
        # 获取模型路径
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
    
    # 选择模式
    mode = rs.GetString("Mode", "Solid", ["Solid", "Wireframe"])
    if not mode:
        return None
    
    create_solid = (mode == "Solid")
    
    try:
        # 删除旧图层
        layer_name = "SAP2000_ByMaterial"
        from PySap2000.visualization.rhino.import_helpers import delete_old_layer, build_elements_by_category
        delete_old_layer(layer_name)
        
        print(f"Importing {mode} model (by material)...")
        
        # 加载 JSON 数据
        from PySap2000.geometry.element_geometry import Model3D
        
        with open(json_file, 'r', encoding='utf-8') as f:
            json_str = f.read()
        
        model_3d = Model3D.from_json(json_str)
        
        # 按材料分组
        material_groups = {}
        for elem in model_3d.elements:
            material = elem.material if elem.material else "Unknown"
            if material not in material_groups:
                material_groups[material] = []
            material_groups[material].append(elem)
        
        # 为每种材料生成随机颜色
        material_colors = {}
        for material in material_groups.keys():
            r = random.randint(50, 230)
            g = random.randint(50, 230)
            b = random.randint(50, 230)
            material_colors[material] = (r, g, b)
        
        all_guids, stats = build_elements_by_category(
            material_groups, material_colors, layer_name, create_solid
        )
        
        rs.MessageBox(
            f"Import complete!\n\n"
            f"Total: {stats['success']} objects\n"
            f"Materials: {len(material_groups)}\n"
            f"Mode: {mode}",
            0,
            "Success"
        )
        
        return all_guids
        
    except Exception as e:
        import traceback
        error_msg = f"Error: {str(e)}\n\n{traceback.format_exc()}"
        print(error_msg)
        rs.MessageBox(error_msg, 16, "Error")
        return None


if __name__ == "__main__":
    import_by_material()
