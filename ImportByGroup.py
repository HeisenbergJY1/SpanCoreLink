# -*- coding: utf-8 -*-
"""
ImportByGroup.py - 按组导入 SAP2000 模型到 Rhino

弹出对话框让用户选择组，然后只导入用户选择的组
每个组分图层，每个组图层分配随机颜色

依赖: 需要先运行 GetSapModel.py 导出模型数据到 JSON 文件
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


def import_by_group():
    """按组导入 SAP2000 模型"""
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
    
    try:
        # 加载 JSON 数据
        from PySap2000.geometry.element_geometry import Model3D
        
        print(f"Loading model data: {json_file}")
        with open(json_file, 'r', encoding='utf-8') as f:
            json_str = f.read()
        
        # 转换为 Model3D 对象
        model_3d = Model3D.from_json(json_str)
        
        # 建立单元名称索引（用于后续按组匹配）
        elem_by_name = {elem.name: elem for elem in model_3d.elements}
        
        # 从 SAP2000 API 获取所有组名（完整列表）
        ret = app.model.GroupDef.GetNameList(0, [])
        if isinstance(ret, (list, tuple)) and len(ret) >= 2:
            all_group_names = list(ret[1]) if ret[1] else []
        else:
            all_group_names = []
        
        # 过滤掉 'ALL' 和系统组（以 ~ 开头）
        user_groups = [g for g in all_group_names if g and g != "ALL" and not g.startswith("~")]
        
        # 过滤掉没有杆件/索单元的空组
        print(f"Filtering {len(user_groups)} user groups...")
        all_groups = []
        for g in user_groups:
            try:
                ret_g = app.model.GroupDef.GetAssignments(g, 0, [], [])
                if isinstance(ret_g, (list, tuple)) and len(ret_g) >= 3:
                    obj_types = list(ret_g[1]) if ret_g[1] else []
                    # type 2=Frame, 3=Cable，只保留包含这些单元的组
                    has_elements = any(t in (2, 3) for t in obj_types)
                    if has_elements:
                        all_groups.append(g)
            except Exception:
                continu
        
        if not all_groups:
            rs.MessageBox("No user groups found in model\n\nTip: Make sure groups are defined", 48, "Info")
            return None
        
        # 排序组名
        sorted_groups = sorted(all_groups)
        
        print(f"Found {len(sorted_groups)} user groups: {sorted_groups}")
        
        # 弹出对话框让用户选择组
        selected_groups = rs.MultiListBox(
            sorted_groups,
            message="Select groups to import (multi-select):",
            title="Select Groups",
            defaults=None
        )
        
        if not selected_groups:
            print("User cancelled selection")
            return None
        
        selected_groups = list(selected_groups)
        print(f"User selected {len(selected_groups)} groups: {selected_groups}")
        
    except Exception as e:
        import traceback
        error_msg = f"Error loading model data: {str(e)}\n\n{traceback.format_exc()}"
        print(error_msg)
        rs.MessageBox(error_msg, 16, "Error")
        return None
    
    # 选择模式
    mode = rs.GetString("Mode", "Solid", ["Solid", "Wireframe"])
    if not mode:
        return None
    
    create_solid = (mode == "Solid")
    
    try:
        # 删除旧图层
        layer_name = "SAP2000_ByGroup"
        from PySap2000.visualization.rhino.import_helpers import delete_old_layer, build_elements_by_category
        delete_old_layer(layer_name)
        
        print(f"Importing {mode} model (by group)...")
        
        # 从 SAP2000 API 查询每个选中组包含的单元
        group_elements = {}
        for group in selected_groups:
            try:
                ret = app.model.GroupDef.GetAssignments(group, 0, [], [])
                if isinstance(ret, (list, tuple)) and len(ret) >= 3:
                    obj_types = list(ret[1]) if ret[1] else []
                    obj_names = list(ret[2]) if ret[2] else []
                    
                    elems = []
                    for j in range(len(obj_types)):
                        if obj_types[j] in (2, 3):
                            name = obj_names[j]
                            if name in elem_by_name:
                                elems.append(elem_by_name[name])
                    
                    if elems:
                        group_elements[group] = elems
            except Exception:
                continue
        
        if not group_elements:
            rs.MessageBox("No elements found in selected groups", 48, "Info")
            return None
        
        # 为每个组生成随机颜色
        group_colors = {}
        for group in group_elements.keys():
            r = random.randint(50, 230)
            g = random.randint(50, 230)
            b = random.randint(50, 230)
            group_colors[group] = (r, g, b)
        
        all_guids, stats = build_elements_by_category(
            group_elements, group_colors, layer_name, create_solid
        )
        
        rs.MessageBox(
            f"Import complete!\n\n"
            f"Total: {stats['success']} objects\n"
            f"Groups: {len(group_elements)}\n"
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
    import_by_group()
