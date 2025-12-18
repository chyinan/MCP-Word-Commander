import logging
import os
import json
from typing import List, Optional, Dict, Any
from mcp.server.fastmcp import FastMCP
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# 初始化 MCP Server
mcp = FastMCP("word-commander")

# 设置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def _get_abs_path(path: str) -> str:
    """获取绝对路径，支持相对路径"""
    if os.path.isabs(path):
        return path
    return os.path.abspath(os.path.join(os.getcwd(), path))

@mcp.tool()
def create_new_document(file_path: str) -> str:
    """
    创建一个新的空白 Word 文档。
    
    Args:
        file_path: 保存文档的路径 (例如: "output.docx")
    """
    try:
        doc = Document()
        abs_path = _get_abs_path(file_path)
        doc.save(abs_path)
        return f"Successfully created new document at {abs_path}"
    except Exception as e:
        return f"Error creating document: {str(e)}"

@mcp.tool()
def get_document_info(file_path: str) -> str:
    """
    获取 Word 文档的基本信息概览，包括段落数、表格数、各段落的简要信息。
    用于在读取完整内容前快速了解文档结构。
    
    Args:
        file_path: 文档路径
        
    Returns:
        JSON 格式的文档概览信息。
    """
    try:
        abs_path = _get_abs_path(file_path)
        if not os.path.exists(abs_path):
            return json.dumps({"error": f"File not found: {abs_path}"}, ensure_ascii=False)
            
        doc = Document(abs_path)
        
        # 统计段落信息
        paragraphs_summary = []
        non_empty_count = 0
        for i, para in enumerate(doc.paragraphs):
            text_preview = para.text.strip()[:50] + "..." if len(para.text.strip()) > 50 else para.text.strip()
            if text_preview:
                non_empty_count += 1
                paragraphs_summary.append({
                    "index": i,
                    "preview": text_preview,
                    "style": para.style.name
                })
        
        info = {
            "file_path": abs_path,
            "total_paragraphs": len(doc.paragraphs),
            "non_empty_paragraphs": non_empty_count,
            "total_tables": len(doc.tables),
            "paragraphs_summary": paragraphs_summary
        }
        
        return json.dumps(info, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


@mcp.tool()
def read_document_structure(
    file_path: str, 
    start_index: int = 0, 
    limit: int = 50,
    include_empty: bool = False
) -> str:
    """
    分段读取 Word 文档的内容和样式信息。支持分页读取大文档。
    
    Args:
        file_path: 文档路径
        start_index: 从第几个段落开始读取 (0-based)
        limit: 最多读取多少个段落 (默认50)
        include_empty: 是否包含空段落 (默认False)
        
    Returns:
        JSON 格式的字符串，包含段落文本和对应的样式信息。
    """
    try:
        abs_path = _get_abs_path(file_path)
        if not os.path.exists(abs_path):
            return json.dumps({"error": f"File not found: {abs_path}"}, ensure_ascii=False)
            
        doc = Document(abs_path)
        content = []
        total = len(doc.paragraphs)
        count = 0
        
        for i, para in enumerate(doc.paragraphs):
            if i < start_index:
                continue
                
            # 跳过空段落（除非指定包含）
            if not include_empty and not para.text.strip():
                continue
            
            if count >= limit:
                break
                
            style_info = {
                "index": i,
                "text": para.text,
                "style_name": para.style.name,
                "alignment": str(para.alignment) if para.alignment else "LEFT",
                "runs": []
            }
            
            # 深入分析 Run (具体的文字片段) 以获取具体字体信息
            for run in para.runs:
                run_info = {
                    "text": run.text,
                    "bold": run.bold,
                    "italic": run.italic,
                    "font_name": run.font.name,
                    "font_size": run.font.size.pt if run.font.size else None
                }
                # 尝试获取中文字体设置 (complex script)
                if run._element.rPr is not None:
                    rFonts = run._element.rPr.find(qn('w:rFonts'))
                    if rFonts is not None:
                        if '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia' in rFonts.attrib:
                            run_info["east_asia_font"] = rFonts.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia']
                
                style_info["runs"].append(run_info)
            
            content.append(style_info)
            count += 1
        
        result = {
            "total_paragraphs": total,
            "returned_count": len(content),
            "start_index": start_index,
            "has_more": (start_index + count) < total,
            "paragraphs": content
        }
            
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)

@mcp.tool()
def read_tables(file_path: str, table_index: Optional[int] = None) -> str:
    """
    读取 Word 文档中的表格内容。
    
    Args:
        file_path: 文档路径
        table_index: 指定读取第几个表格 (0-based)，不指定则读取所有表格
        
    Returns:
        JSON 格式的表格数据。
    """
    try:
        abs_path = _get_abs_path(file_path)
        if not os.path.exists(abs_path):
            return json.dumps({"error": f"File not found: {abs_path}"}, ensure_ascii=False)
            
        doc = Document(abs_path)
        tables_data = []
        
        for idx, table in enumerate(doc.tables):
            if table_index is not None and idx != table_index:
                continue
                
            table_content = {
                "table_index": idx,
                "rows": len(table.rows),
                "cols": len(table.columns),
                "data": []
            }
            
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    row_data.append(cell.text)
                table_content["data"].append(row_data)
            
            tables_data.append(table_content)
        
        return json.dumps({
            "total_tables": len(doc.tables),
            "tables": tables_data
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


def _apply_paragraph_format(paragraph, run, font_name: str, font_size: float, 
                            is_bold: bool, alignment: str, indent_first_line: float,
                            line_spacing: Optional[float] = None):
    """内部函数：应用段落格式"""
    align_map = {
        "LEFT": WD_ALIGN_PARAGRAPH.LEFT,
        "CENTER": WD_ALIGN_PARAGRAPH.CENTER,
        "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT,
        "JUSTIFY": WD_ALIGN_PARAGRAPH.JUSTIFY
    }
    paragraph.alignment = align_map.get(alignment.upper(), WD_ALIGN_PARAGRAPH.LEFT)
    
    if indent_first_line > 0:
        paragraph.paragraph_format.first_line_indent = Pt(font_size * indent_first_line)
    
    if line_spacing:
        paragraph.paragraph_format.line_spacing = Pt(line_spacing)
    
    run.bold = is_bold
    run.font.size = Pt(font_size)
    run.font.name = font_name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), font_name)


@mcp.tool()
def add_formatted_paragraph(
    file_path: str, 
    text: str, 
    font_name: str = "SimSun", 
    font_size: float = 12.0, 
    is_bold: bool = False,
    alignment: str = "LEFT",
    indent_first_line: float = 0.0,
    line_spacing: Optional[float] = None
) -> str:
    """
    向文档追加带有特定样式的段落。支持设置中文字体（如宋体/黑体）。
    
    Args:
        file_path: 文档路径
        text: 段落内容
        font_name: 字体名称 (默认 "SimSun" 即宋体, 可选 "Microsoft YaHei", "Times New Roman" 等)
        font_size: 字号 (pt)
        is_bold: 是否加粗
        alignment: 对齐方式 ("LEFT", "CENTER", "RIGHT", "JUSTIFY")
        indent_first_line: 首行缩进字符数 (例如 2.0 代表缩进两个字符)
        line_spacing: 行距 (pt)，例如 20.0 代表固定值20磅
    """
    try:
        abs_path = _get_abs_path(file_path)
        if not os.path.exists(abs_path):
            return json.dumps({"error": f"File not found: {abs_path}"}, ensure_ascii=False)
            
        doc = Document(abs_path)
        paragraph = doc.add_paragraph()
        run = paragraph.add_run(text)
        
        _apply_paragraph_format(paragraph, run, font_name, font_size, 
                               is_bold, alignment, indent_first_line, line_spacing)
        
        doc.save(abs_path)
        return json.dumps({"success": True, "message": "Successfully added formatted paragraph."}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


@mcp.tool()
def replace_paragraph(
    file_path: str,
    paragraph_index: int,
    new_text: str,
    font_name: str = "SimSun",
    font_size: float = 12.0,
    is_bold: bool = False,
    alignment: str = "LEFT",
    indent_first_line: float = 0.0,
    line_spacing: Optional[float] = None
) -> str:
    """
    替换指定段落的内容，并应用新的格式。
    
    Args:
        file_path: 文档路径
        paragraph_index: 要替换的段落索引 (0-based)
        new_text: 新的段落内容
        font_name: 字体名称
        font_size: 字号 (pt)
        is_bold: 是否加粗
        alignment: 对齐方式
        indent_first_line: 首行缩进字符数
        line_spacing: 行距 (pt)
    """
    try:
        abs_path = _get_abs_path(file_path)
        if not os.path.exists(abs_path):
            return json.dumps({"error": f"File not found: {abs_path}"}, ensure_ascii=False)
            
        doc = Document(abs_path)
        
        if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
            return json.dumps({
                "error": f"Invalid paragraph index: {paragraph_index}. Document has {len(doc.paragraphs)} paragraphs."
            }, ensure_ascii=False)
        
        para = doc.paragraphs[paragraph_index]
        
        # 清除原有内容
        para.clear()
        
        # 添加新内容
        run = para.add_run(new_text)
        _apply_paragraph_format(para, run, font_name, font_size, 
                               is_bold, alignment, indent_first_line, line_spacing)
        
        doc.save(abs_path)
        return json.dumps({
            "success": True, 
            "message": f"Successfully replaced paragraph at index {paragraph_index}."
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


@mcp.tool()
def insert_paragraph_after(
    file_path: str,
    after_index: int,
    text: str,
    font_name: str = "SimSun",
    font_size: float = 12.0,
    is_bold: bool = False,
    alignment: str = "LEFT",
    indent_first_line: float = 0.0,
    line_spacing: Optional[float] = None
) -> str:
    """
    在指定段落之后插入新段落。
    
    Args:
        file_path: 文档路径
        after_index: 在此段落索引之后插入 (0-based)
        text: 新段落内容
        font_name: 字体名称
        font_size: 字号 (pt)
        is_bold: 是否加粗
        alignment: 对齐方式
        indent_first_line: 首行缩进字符数
        line_spacing: 行距 (pt)
    """
    try:
        abs_path = _get_abs_path(file_path)
        if not os.path.exists(abs_path):
            return json.dumps({"error": f"File not found: {abs_path}"}, ensure_ascii=False)
            
        doc = Document(abs_path)
        
        if after_index < 0 or after_index >= len(doc.paragraphs):
            return json.dumps({
                "error": f"Invalid paragraph index: {after_index}. Document has {len(doc.paragraphs)} paragraphs."
            }, ensure_ascii=False)
        
        # 获取目标段落
        target_para = doc.paragraphs[after_index]
        
        # 在目标段落之后插入新段落
        new_para = OxmlElement('w:p')
        target_para._element.addnext(new_para)
        
        # 创建新的段落对象
        from docx.text.paragraph import Paragraph
        new_paragraph = Paragraph(new_para, target_para._parent)
        
        run = new_paragraph.add_run(text)
        _apply_paragraph_format(new_paragraph, run, font_name, font_size,
                               is_bold, alignment, indent_first_line, line_spacing)
        
        doc.save(abs_path)
        return json.dumps({
            "success": True,
            "message": f"Successfully inserted paragraph after index {after_index}."
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


@mcp.tool()
def search_and_replace(
    file_path: str,
    search_text: str,
    replace_text: str,
    match_case: bool = True
) -> str:
    """
    在文档中搜索并替换文本（保持原有格式）。
    
    Args:
        file_path: 文档路径
        search_text: 要搜索的文本
        replace_text: 替换为的文本
        match_case: 是否区分大小写 (默认True)
    """
    try:
        abs_path = _get_abs_path(file_path)
        if not os.path.exists(abs_path):
            return json.dumps({"error": f"File not found: {abs_path}"}, ensure_ascii=False)
            
        doc = Document(abs_path)
        count = 0
        
        for para in doc.paragraphs:
            for run in para.runs:
                if match_case:
                    if search_text in run.text:
                        run.text = run.text.replace(search_text, replace_text)
                        count += 1
                else:
                    import re
                    if re.search(re.escape(search_text), run.text, re.IGNORECASE):
                        run.text = re.sub(re.escape(search_text), replace_text, run.text, flags=re.IGNORECASE)
                        count += 1
        
        # 同时处理表格中的文本
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if match_case:
                                if search_text in run.text:
                                    run.text = run.text.replace(search_text, replace_text)
                                    count += 1
                            else:
                                import re
                                if re.search(re.escape(search_text), run.text, re.IGNORECASE):
                                    run.text = re.sub(re.escape(search_text), replace_text, run.text, flags=re.IGNORECASE)
                                    count += 1
        
        doc.save(abs_path)
        return json.dumps({
            "success": True,
            "message": f"Replaced {count} occurrence(s) of '{search_text}'."
        }, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)

@mcp.tool()
def create_table_with_data(
    file_path: str,
    rows: int,
    cols: int,
    data: List[List[str]],
    header_bold: bool = True
) -> str:
    """
    在文档末尾创建一个表格并填充数据。
    
    Args:
        file_path: 文档路径
        rows: 行数
        cols: 列数
        data: 二维数组，包含要填充的数据 [['Header1', 'Header2'], ['Val1', 'Val2']]
        header_bold: 第一行是否加粗
    """
    try:
        abs_path = _get_abs_path(file_path)
        if not os.path.exists(abs_path):
            return f"File not found: {abs_path}"
            
        doc = Document(abs_path)
        table = doc.add_table(rows=rows, cols=cols)
        table.style = 'Table Grid' # 使用默认带边框的样式
        
        # 填充数据
        for r in range(min(rows, len(data))):
            row_data = data[r]
            for c in range(min(cols, len(row_data))):
                cell = table.cell(r, c)
                cell.text = str(row_data[c])
                
                # 如果是表头且需要加粗
                if r == 0 and header_bold:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.bold = True
                            
        doc.save(abs_path)
        return "Successfully created table."
    except Exception as e:
        return f"Error creating table: {str(e)}"

if __name__ == "__main__":
    mcp.run()

