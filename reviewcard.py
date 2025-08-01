from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

def sync_table_content(doc_path, output_path):
    doc = Document(doc_path)
    
    if not doc.tables:
        print("文档中没有表格")
        return
    
    # 从第一个表格提取同步内容
    sync_data = {}
    first_table = doc.tables[0]
    
    # 示例：提取项目编号和工程名称
    sync_data["project_no"] = first_table.cell(0, 1).text
    sync_data["project_name"] = first_table.cell(1, 1).text
    sync_data["project_manager"] = first_table.cell(1, 3).text
    
    # 同步到其他表格
    for table in doc.tables[1:]:
        for row in table.rows:
            for cell in row.cells:
                # 匹配占位符并替换
                if "项目名称" in cell.text:
                    cell.text = cell.text.replace("项目名称", f"项目名称: {sync_data['project_name']}")
                if "项目编号" in cell.text:
                    cell.text = cell.text.replace("项目编号", f"项目编号: {sync_data['project_no']}")
                if "项目经理" in cell.text:
                    cell.text = cell.text.replace("项目经理", f"项目经理: {sync_data['project_manager']}")
                
                # 保留原始格式
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10)  # 保持字体大小
                        run.font.name = '宋体'  # 保持中文字体
    
    doc.save(output_path)
    print(f"同步完成! 已保存到: {output_path}")

if __name__ == "__main__":
    input_file = "校审卡.docx"
    output_file = "校审卡_同步后.docx"
    sync_table_content(input_file, output_file)