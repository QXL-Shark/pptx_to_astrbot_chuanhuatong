import os
import zipfile
import shutil
import xml.etree.ElementTree as ET
from pathlib import Path

def parse_presentation_xml(xml_file):
    """
    解析presentation.xml文件，提取章节和幻灯片顺序信息
    """
    # 定义XML命名空间
    namespaces = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    }
    
    # 解析XML文件
    tree = ET.parse(xml_file)
    root = tree.getroot()
    
    # 获取所有幻灯片的ID和顺序
    sld_id_lst = root.find('.//p:sldIdLst', namespaces)
    slide_order = []
    if sld_id_lst is not None:
        for sld_id in sld_id_lst.findall('p:sldId', namespaces):
            slide_id = sld_id.get('id')
            if slide_id:
                slide_order.append(slide_id)
    
    # 获取章节信息
    sections = []
    section_lst = root.find('.//p14:sectionLst', namespaces)
    
    if section_lst is not None:
        for section in section_lst.findall('p14:section', namespaces):
            section_name = section.get('name')
            sld_id_lst = section.find('p14:sldIdLst', namespaces)
            
            if section_name and sld_id_lst:
                slide_ids_in_section = []
                for sld_id in sld_id_lst.findall('p14:sldId', namespaces):
                    slide_id = sld_id.get('id')
                    if slide_id:
                        slide_ids_in_section.append(slide_id)
                
                # 将幻灯片ID转换为顺序索引（从1开始）
                ppt_indices = []
                for slide_id in slide_ids_in_section:
                    if slide_id in slide_order:
                        ppt_indices.append(slide_order.index(slide_id) + 1)
                
                sections.append({
                    "section_name": section_name,
                    "ppt_index": ppt_indices
                })
    
    return sections

def reorganize_ppt_structure(root_dir, ppt_name):
    """
    重组PPT结构
    """
    pptx_file = os.path.join(root_dir, f"{ppt_name}.pptx")
    png_dir = os.path.join(root_dir, ppt_name)
    extract_dir = os.path.join(root_dir, "extract")
    restructured_dir = os.path.join(root_dir, "restructured")
    
    # 1. 解压PPTX文件
    print(f"解压PPTX文件: {pptx_file}")
    with zipfile.ZipFile(pptx_file, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)
    
    # 2. 解析presentation.xml
    presentation_xml = os.path.join(extract_dir, "ppt", "presentation.xml")
    print(f"解析XML文件: {presentation_xml}")
    sections = parse_presentation_xml(presentation_xml)
    
    print("解析到的章节信息:")
    for section in sections:
        print(f"  章节: {section['section_name']}, 幻灯片索引: {section['ppt_index']}")
    
    # 3. 创建重组目录结构并移动PNG文件
    print("创建重组目录结构...")
    os.makedirs(restructured_dir, exist_ok=True)
    
    for section in sections:
        section_dir = os.path.join(restructured_dir, section["section_name"])
        os.makedirs(section_dir, exist_ok=True)
        
        print(f"处理章节: {section['section_name']}")
        for ppt_index in section["ppt_index"]:
            png_filename = f"幻灯片{ppt_index}.png"
            source_png = os.path.join(png_dir, png_filename)
            dest_png = os.path.join(section_dir, png_filename)
            
            if os.path.exists(source_png):
                shutil.copy2(source_png, dest_png)
                print(f"  复制: {png_filename} -> {section['section_name']}/")
            else:
                print(f"  警告: 文件不存在 - {source_png}")
    
    # 4. 清理临时解压文件
    print("清理临时文件...")
    shutil.rmtree(extract_dir)
    
    print(f"重组完成！结果保存在: {restructured_dir}")
    return sections

def get_pptx_filename(folder_path):
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.pptx') and not filename.startswith('~$'):
            return os.path.splitext(filename)[0]
    return None

def process_dir(root_dir):
    """
    主函数
    """
    # 配置参数
    ppt_name = get_pptx_filename(root_dir)
    if not ppt_name:
        print("错误: 未找到PPTX文件")
        return
    
    # 验证输入
    pptx_path = os.path.join(root_dir, f"{ppt_name}.pptx")
    png_dir_path = os.path.join(root_dir, ppt_name)
    
    if not os.path.exists(pptx_path):
        print(f"错误: PPTX文件不存在 - {pptx_path}")
        return
    
    if not os.path.exists(png_dir_path):
        print(f"错误: PNG文件夹不存在 - {png_dir_path}")
        return
    
    # 执行重组
    try:
        sections = reorganize_ppt_structure(root_dir, ppt_name)
        
        # 显示最终结果
        print("\n最终章节结构:")
        for i, section in enumerate(sections):
            print(f"{i+1}. 章节: {section['section_name']}")
            print(f"   包含幻灯片: {section['ppt_index']}")
            
    except Exception as e:
        print(f"处理过程中发生错误: {e}")
        import traceback
        traceback.print_exc()

def get_sibling_folders():
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 获取同级目录下的所有条目
    all_items = os.listdir(current_dir)
    
    # 筛选出文件夹并获取绝对路径
    folders = []
    for item in all_items:
        item_path = os.path.join(current_dir, item)
        if os.path.isdir(item_path):
            folders.append(os.path.abspath(item_path))
    
    return folders

def main():
    dirs=get_sibling_folders()
    for dir in dirs:
        process_dir(dir)

if __name__ == '__main__':
    main()
