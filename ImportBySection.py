# -*- coding: utf-8 -*-
"""
ImportBySection.py - 按截面导入 SAP2000 模型到 Rhino

按截面分图层，每个截面图层分配随机颜色
支持使用原始截面名或标准化截面名作为图层名
使用标准化名称时，相同标准化名称的不同截面会合并到同一图层

依赖: 需要先运行 GetSapModel.py 导出模型数据到 JSON 文件
"""
import os
import sys

# ============ 开发环境配置 ============
DEV_MODE = False

if DEV_MODE:
    _script_dir = os.path.dirname(os.path.abspath(__file__))
    _workspace_root = os.path.dirname(_script_dir)
    if _workspace_root not in sys.path:
        sys.path.insert(0, _workspace_root)
    print(f"[DEV] Using local development version: {_workspace_root}")
# =====================================

import rhinoscriptsyntax as rs
import random


def import_by_section():
    """按截面导入 SAP2000 模型"""
    try:
        from PySap2000.application import Application
        
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
    
    # 选择图层命名方式
    naming = rs.GetString("Layer naming", "Standard", ["Standard", "Original"])
    if not naming:
        return None
    use_standard_name = (naming == "Standard")
    
    # 选择模式
    mode = rs.GetString("Mode", "Solid", ["Solid", "Wireframe"])
    if not mode:
        return None
    create_solid = (mode == "Solid")
    
    try:
        # 加载 JSON 数据
        from PySap2000.geometry.element_geometry import Model3D
        
        print(f"Loading model data...")
        with open(json_file, 'r', encoding='utf-8') as f:
            json_str = f.read()
        
        model_3d = Model3D.from_json(json_str)
        
        if not model_3d.elements:
            rs.MessageBox("No elements found in model", 48, "Info")
            return None
        
        # 收集所有截面名
        all_sections = set()
        for elem in model_3d.elements:
            if hasattr(elem, 'section_name') and elem.section_name:
                all_sections.add(elem.section_name)
        
        print(f"Found {len(all_sections)} unique sections")
        
        # 建立截面名到显示名的映射
        # 如果使用标准化名称，相同标准化名称的截面会映射到同一个显示名
        section_to_display = {}  # {原始截面名: 显示名}
        
        if use_standard_name:
            from PySap2000.global_parameters.units import Units, UnitSystem
            from PySap2000.section.frame_section import FrameSection
            from PySap2000.section.cable_section import CableSection
            from PySap2000.geometry.element_geometry import CableElement3D
            
            # 切换到 N-mm-C 单位（标准化名称需要 mm 单位）
            original_units = Units.get_present_units(app.model)
            Units.set_present_units(app.model, UnitSystem.N_MM_C)
            
            try:
                print("Generating standard section names...")
                
                # 判断截面是 Frame 还是 Cable 类型
                section_types = {}  # {截面名: "frame" | "cable"}
                for elem in model_3d.elements:
                    if hasattr(elem, 'section_name') and elem.section_name:
                        if isinstance(elem, CableElement3D):
                            section_types[elem.section_name] = "cable"
                        else:
                            section_types[elem.section_name] = "frame"
                
                for section in all_sections:
                    try:
                        if section_types.get(section) == "cable":
                            # 索截面
                            cable_sec = CableSection.get_by_name(app.model, section)
                            std_name = cable_sec.standard_name
                            print(f"  [Cable] {section}: area={cable_sec.area:.2f} -> {std_name}")
                        else:
                            # 杆件截面
                            frame_sec = FrameSection.get_by_name(app.model, section)
                            std_name = frame_sec.standard_name
                            if std_name != section:
                                print(f"  [Frame] {section} -> {std_name}")
                        
                        section_to_display[section] = std_name
                    except Exception as e:
                        # 获取失败，使用原名称
                        print(f"  [Error] {section}: {e}")
                        section_to_display[section] = section
            finally:
                # 确保无论如何都恢复单位
                Units.set_present_units(app.model, original_units)
        else:
            for section in all_sections:
                section_to_display[section] = section
        
        # 按显示名分组（相同标准化名称的截面会合并）
        display_elements = {}  # {显示名: [元素列表]}
        for elem in model_3d.elements:
            if hasattr(elem, 'section_name') and elem.section_name:
                display_name = section_to_display.get(elem.section_name, elem.section_name)
            else:
                display_name = "Unknown"
            
            if display_name not in display_elements:
                display_elements[display_name] = []
            display_elements[display_name].append(elem)
        
        # 删除旧图层
        layer_name = "SAP2000_BySection"
        from PySap2000.visualization.rhino.import_helpers import delete_old_layer, build_elements_by_category
        delete_old_layer(layer_name)
        
        # 为每个显示名生成随机颜色
        display_colors = {}
        for name in display_elements.keys():
            r = random.randint(50, 230)
            g = random.randint(50, 230)
            b = random.randint(50, 230)
            display_colors[name] = (r, g, b)
        
        all_guids, stats = build_elements_by_category(
            display_elements, display_colors, layer_name, create_solid
        )
        
        rs.MessageBox(
            f"Import complete!\n\n"
            f"Total: {stats['success']} objects\n"
            f"Sections: {len(display_elements)}\n"
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
    import_by_section()
