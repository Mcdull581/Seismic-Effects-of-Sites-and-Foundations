# -*- coding: utf-8 -*-
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

"""
build_seismic_report.py
功能：
  1. 解析《全国抗震设防烈度表.txt》，提取全国所有地级市/区县的抗震设防数据
  2. 合并到 seismic_evaluation_data.json 中，生成完整版 seismic_evaluation_data_full.json
  3. 基于完整数据，生成排版严密的 Word 规范报告
"""

import json
import re
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ============================================================
# 常量
# ============================================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT_JSON = os.path.join(BASE_DIR, "data", "seismic_data_full.json")
INPUT_TXT  = os.path.join(BASE_DIR, "data", "全国抗震设防烈度表.txt")
OUTPUT_JSON = os.path.join(BASE_DIR, "data", "seismic_data_full.json")
OUTPUT_DOCX = os.path.join(BASE_DIR, "output", "场地和地基地震效应评价规范.docx")


# ============================================================
# 模块一：解析全国抗震设防烈度表.txt
# ============================================================
def parse_national_seismic_txt(filepath):
    """
    解析全国抗震设防烈度表.txt，返回结构化列表。
    每条记录: {province, city, intensity, pga, group}
    """
    with open(filepath, "r", encoding="utf-8") as f:
        text = f.read()

    # 清洗OCR/排版错误
    text = text.replace("0..10g", "0.10g")
    text = text.replace("0..1Og", "0.10g")
    text = text.replace("0.1Og", "0.10g")
    text = text.replace("0.lOg", "0.10g")
    text = text.replace("0.l0g", "0.10g")

    records = []
    current_province = ""

    # 按行处理
    lines = text.split("\n")

    # 正则：匹配省份标题，如 "A.0.2 河北省"
    re_province = re.compile(r"A\.0\.\d+\s+(.+?省|.+?市|.+?自治区|.+?特区|台湾省|港澳特区和台湾省)")
    # 正则：匹配烈度行，如 "1. 抗震设防烈度为8度，设计基本地震加速度值为0.20g："
    re_intensity = re.compile(
        r"(?:\d+\s+)?抗震设防烈度(?:不低于|不小于|为)?\s*[≥>]?\s*(\d+)度[，,]设计基本地震加速度值(?:不小于|不低于|为)?\s*[≥>]?\s*(0\.\d+)g"
    )
    # 正则：匹配分组行，如 "第一组：北京（东城...），延庆；"
    re_group = re.compile(r"(第[一二三]组)[：:]\s*(.+)")

    # 用于暂存当前烈度/加速度
    cur_intensity = None
    cur_pga = None

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # 检测省份
        m_prov = re_province.search(line)
        if m_prov:
            current_province = m_prov.group(1).strip()
            continue

        # 检测烈度/加速度行
        m_int = re_intensity.search(line)
        if m_int:
            cur_intensity = int(m_int.group(1))
            cur_pga = float(m_int.group(2))
            continue

        # 检测分组行
        m_grp = re_group.search(line)
        if m_grp and cur_intensity is not None:
            group_name = m_grp.group(1)  # 第一组/第二组/第三组
            cities_text = m_grp.group(2)

            # 解析城市列表：按 ；或 ，分割，但需要处理括号内的内容
            # 策略：先按中文分号分割段落，再按中文逗号分割城市
            segments = re.split(r"[；;]", cities_text)
            for seg in segments:
                seg = seg.strip().rstrip("。")
                if not seg:
                    continue
                # 提取带括号的城市和不带括号的城市
                # 如 "北京（东城、西城），延庆" -> ["北京（东城、西城）", "延庆"]
                # 用正则匹配：城市名（可选括号内容）
                parts = re.findall(r"[一-鿿]+(?:（[^）]*）)?", seg)
                for part in parts:
                    part = part.strip()
                    if not part:
                        continue
                    # 如果有括号，拆分为多个区县
                    m_paren = re.match(r"^([一-鿿]+)（(.+)）$", part)
                    if m_paren:
                        parent_city = m_paren.group(1)
                        districts = re.split(r"[、,，]", m_paren.group(2))
                        for dist in districts:
                            dist = dist.strip()
                            if dist:
                                records.append({
                                    "province": current_province,
                                    "city": f"{parent_city}({dist})",
                                    "intensity": cur_intensity,
                                    "pga": cur_pga,
                                    "group": group_name
                                })
                    else:
                        records.append({
                            "province": current_province,
                            "city": part,
                            "intensity": cur_intensity,
                            "pga": cur_pga,
                            "group": group_name
                        })

    return records


# ============================================================
# 模块二：数据合并
# ============================================================
def merge_data(json_path, txt_path, output_path):
    """读取JSON和TXT，合并后保存完整版JSON"""
    # 读取原始JSON
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # 解析全国数据
    national_records = parse_national_seismic_txt(txt_path)
    print(f"[数据合并] 从TXT解析出 {len(national_records)} 条全国抗震设防记录")

    # 转换为JSON中的格式
    appendix_a_data = []
    for rec in national_records:
        group_num = rec["group"].replace("第", "").replace("组", "")
        group_map = {"一": "第一组", "二": "第二组", "三": "第三组"}
        appendix_a_data.append({
            "province": rec["province"],
            "city": rec["city"],
            "intensity": rec["intensity"],
            "pga": rec["pga"],
            "group": group_map.get(group_num, rec["group"])
        })

    # 替换附录A数据
    data["appendix_a_seismic_intensity"]["data"] = appendix_a_data
    data["appendix_a_seismic_intensity"]["description"] = (
        "《建筑抗震设计规范》附录A - 全国主要城镇抗震设防烈度（完整数据）"
    )
    data["appendix_a_seismic_intensity"]["note_demo_data"] = (
        f"已补全全国数据，共 {len(appendix_a_data)} 条记录"
    )

    # 附录C同样替换为全国数据（使用相同数据源，字段略有不同）
    appendix_c_data = []
    for rec in national_records:
        tg_map = {"第一组": 0.35, "第二组": 0.40, "第三组": 0.45}
        tg_group_map = {"第一组": "第一组", "第二组": "第二组", "第三组": "第三组"}
        appendix_c_data.append({
            "area": rec["city"],
            "province": rec["province"],
            "pga": rec["pga"],
            "tg_group": tg_group_map.get(rec["group"], "第二组"),
            "tg_value": tg_map.get(rec["group"], 0.40)
        })

    data["appendix_c_ground_motion"]["data"] = appendix_c_data
    data["appendix_c_ground_motion"]["description"] = (
        "《中国地震动参数区划图》GB 18306-2015 附录C - 地震动参数（完整数据）"
    )
    data["appendix_c_ground_motion"]["note"] = (
        f"已补全全国数据，共 {len(appendix_c_data)} 条记录"
    )

    # 自动补全检测：确保所有关键表都存在且完整
    required_tables = [
        "table_4_1_3_soil_type",
        "table_4_1_6_site_classification",
        "table_4_3_3_liquefaction_depth",
        "table_4_3_4_N0",
        "table_4_3_4_beta",
        "table_4_3_5_liquefaction_grade",
        "table_4_3_6_liquefaction_measures",
        "table_4_1_1_site_classification_type"
    ]
    for tbl_name in required_tables:
        if tbl_name not in data:
            print(f"[警告] 缺失数据表: {tbl_name}")
        elif "data" not in data[tbl_name] or not data[tbl_name]["data"]:
            print(f"[警告] 数据表为空: {tbl_name}")
        else:
            print(f"[OK] 数据表完整: {tbl_name} ({len(data[tbl_name]['data'])} 条)")

    # 保存完整版JSON
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[数据合并] 完整版JSON已保存: {output_path}")

    return data


# ============================================================
# 模块三：Word文档生成
# ============================================================

def set_cell_border(cell, **kwargs):
    """设置单元格边框"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}>'
                          f'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                          f'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                          f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                          f'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                          f'</w:tcBorders>')
    tcPr.append(tcBorders)


def add_bordered_table(doc, headers, rows):
    """添加带边框的表格，表头加粗"""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 写表头
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(str(header))
        run.bold = True
        run.font.size = Pt(10)
        run.font.name = "宋体"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
        set_cell_border(cell)

    # 写数据行
    for r_idx, row_data in enumerate(rows):
        for c_idx, val in enumerate(row_data):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(str(val))
            run.font.size = Pt(10)
            run.font.name = "宋体"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
            set_cell_border(cell)

    return table


def add_step_paragraph(doc, label, content):
    """添加 a/b/c/d/e 格式的段落"""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)

    run_label = p.add_run(f"{label}：")
    run_label.bold = True
    run_label.font.size = Pt(11)
    run_label.font.name = "宋体"
    run_label._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    run_content = p.add_run(str(content))
    run_content.font.size = Pt(11)
    run_content.font.name = "宋体"
    run_content._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")


def add_heading(doc, text, level=1):
    """添加标题"""
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.name = "黑体"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")


def set_doc_style(doc):
    """设置文档默认样式"""
    style = doc.styles["Normal"]
    font = style.font
    font.name = "宋体"
    font.size = Pt(11)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    # 设置页边距
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.18)
        section.right_margin = Cm(3.18)


def generate_docx(data, output_path):
    """根据JSON数据生成Word文档"""
    doc = Document()
    set_doc_style(doc)

    # === 文档标题 ===
    title = doc.add_heading("场地和地基地震效应评价规范", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.name = "黑体"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")

    # 副标题
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = sub.add_run("依据《建筑抗震设计规范》GB 50011-2010 及《中国地震动参数区划图》GB 18306-2015")
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(100, 100, 100)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    doc.add_paragraph("")  # 空行

    # === 七大评价步骤的完整逻辑 ===
    # 由于原始JSON是按数据表组织的，这里需要按文档逻辑重新组织七大步骤
    steps = build_seven_steps(data)

    for step_data in steps:
        # 大标题
        add_heading(doc, step_data["title"], level=1)

        if "sub_steps" in step_data:
            for sub_step in step_data["sub_steps"]:
                # 步骤标题
                if "step_title" in sub_step:
                    p = doc.add_paragraph()
                    run = p.add_run(sub_step["step_title"])
                    run.bold = True
                    run.font.size = Pt(12)
                    run.font.name = "黑体"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")

                # 写入 a/b/c/d/e
                for key in ["input", "basis", "logic", "output", "next"]:
                    if key in sub_step:
                        labels = {"input": "a.\"输入\"", "basis": "b.\"判断依据\"",
                                  "logic": "c.\"判断逻辑\"", "output": "d.\"输出\"",
                                  "next": "e.\"下一步\""}
                        add_step_paragraph(doc, labels[key], sub_step[key])

                # 如果有表格数据，插入表格
                if "table" in sub_step:
                    tbl = sub_step["table"]
                    add_bordered_table(doc, tbl["headers"], tbl["rows"])
                    doc.add_paragraph("")  # 表后空行

        # 步骤间空行
        doc.add_paragraph("")

    # === 附录：核心数据表 ===
    add_heading(doc, "附录：核心规范数据表", level=1)

    # 表4.1.3
    add_data_table_section(doc, data, "table_4_1_3_soil_type", "表4.1.3 土的类型与剪切波速范围")
    # 表4.1.6
    add_data_table_section(doc, data, "table_4_1_6_site_classification", "表4.1.6 场地类别划分")
    # 表4.3.3
    add_data_table_section(doc, data, "table_4_3_3_liquefaction_depth", "表4.3.3 液化土特征深度(m)")
    # 表4.3.4 N0
    add_data_table_section(doc, data, "table_4_3_4_N0", "表4.3.4 液化判别标准贯入锤击数基准值N₀")
    # 表4.3.4 beta
    add_data_table_section(doc, data, "table_4_3_4_beta", "表4.3.4 调整系数β")
    # 表4.3.5
    add_data_table_section(doc, data, "table_4_3_5_liquefaction_grade", "表4.3.5 液化等级")
    # 表4.3.6
    add_data_table_section(doc, data, "table_4_3_6_liquefaction_measures", "表4.3.6 抗液化措施")

    # === 附录：全国抗震设防数据统计 ===
    add_heading(doc, "附录：全国抗震设防数据统计", level=1)

    # 按省份统计
    province_stats = {}
    for rec in data["appendix_a_seismic_intensity"]["data"]:
        prov = rec.get("province", "未知")
        if prov not in province_stats:
            province_stats[prov] = {"count": 0, "intensities": {}}
        province_stats[prov]["count"] += 1
        int_key = str(rec["intensity"])
        province_stats[prov]["intensities"][int_key] = province_stats[prov]["intensities"].get(int_key, 0) + 1

    p = doc.add_paragraph()
    run = p.add_run(f"共覆盖 {len(province_stats)} 个省级行政区，"
                    f"{len(data['appendix_a_seismic_intensity']['data'])} 条记录。")
    run.font.size = Pt(11)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    # 各省统计表
    stat_headers = ["省份", "记录总数", "6度", "7度(0.10g)", "7度(0.15g)", "8度(0.20g)", "8度(0.30g)", "9度"]
    stat_rows = []
    for prov in sorted(province_stats.keys()):
        s = province_stats[prov]
        intensities = s["intensities"]
        # 按PGA细分
        pga_count = {}
        for rec in data["appendix_a_seismic_intensity"]["data"]:
            if rec.get("province") == prov:
                key = f"{rec['intensity']}_{rec['pga']}"
                pga_count[key] = pga_count.get(key, 0) + 1
        stat_rows.append([
            prov, s["count"],
            pga_count.get("6_0.05", 0),
            pga_count.get("7_0.1", 0),
            pga_count.get("7_0.15", 0),
            pga_count.get("8_0.2", 0),
            pga_count.get("8_0.3", 0),
            pga_count.get("9_0.4", 0),
        ])

    add_bordered_table(doc, stat_headers, stat_rows)

    # 保存
    doc.save(output_path)
    print(f"[文档生成] Word报告已保存: {output_path}")


def add_data_table_section(doc, data, table_key, title):
    """添加数据表章节"""
    if table_key not in data:
        return
    tbl_data = data[table_key]
    add_heading(doc, title, level=2)

    # 描述
    if "description" in tbl_data:
        p = doc.add_paragraph()
        run = p.add_run(tbl_data["description"])
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(100, 100, 100)
        run.font.name = "宋体"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    # 构建表格
    if "columns" in tbl_data and "data" in tbl_data:
        columns = tbl_data["columns"]
        rows_data = tbl_data["data"]

        table_rows = []
        for row in rows_data:
            if isinstance(row, dict):
                table_rows.append([str(row.get(c, "")) for c in columns])
            else:
                table_rows.append([str(v) for v in row])

        add_bordered_table(doc, columns, table_rows)
        doc.add_paragraph("")


def build_seven_steps(data):
    """构建七大评价步骤的完整逻辑数据"""
    steps = []

    # ===== 步骤一：场地抗震设防烈度 =====
    steps.append({
        "title": "（一）场地抗震设防烈度",
        "sub_steps": [{
            "step_title": "场地抗震设防烈度判定",
            "input": "提取勘察报告2.1.1地理位置（区、县、镇或街道办）",
            "basis": "《建筑抗震设计规范》附录A、《中国地震动参数区划图》附录C",
            "logic": "根据项目所在的区、县，查询《建筑抗震设计规范》附录A，获得抗震设防烈度和设计地震分组；根据项目所在的镇或街道办，查询《中国地震动参数区划图》附录C，获得二类场地条件下基本地震动峰值加速度和反应谱特征周期。项目跨多区县、乡镇的，分别查询。",
            "output": "场地抗震设防烈度和设计地震分组；二类场地条件下基本地震动峰值加速度和反应谱特征周期。",
            "next": "无"
        }]
    })

    # ===== 步骤二：场地类别划分 =====
    steps.append({
        "title": "（二）场地类别划分",
        "sub_steps": [
            {
                "step_title": "步骤1（根据波速测试报告直接划分）",
                "input": "波速测试报告、编录信息表",
                "basis": "《建筑抗震设计规范》第4.1.6条",
                "logic": "若有波速测试报告，直接获取各钻孔土层等效剪切波速和覆盖层厚度，查询《建筑抗震设计规范》表4.1.6，得到场地类别；若没有波速测试报告，进入步骤2。",
                "output": "场地类别。",
                "next": "没有波速测试报告，进入步骤2。",
                "table": {
                    "headers": ["等效剪切波速(m/s)", "I₀", "I₁", "II", "III", "IV"],
                    "rows": [
                        ["vs > 800", "0", "—", "—", "—", "—"],
                        ["800 ≥ vs > 500", "—", "0", "—", "—", "—"],
                        ["500 ≥ vse > 250", "—", "<5m", "≥5m", "—", "—"],
                        ["250 ≥ vse > 150", "—", "<3m", "3~50m", ">50m", "—"],
                        ["vse ≤ 150", "—", "<3m", "3~15m", "15~80m", ">80m"]
                    ]
                }
            },
            {
                "step_title": "步骤2（计算覆盖层厚度）",
                "input": "编录信息表",
                "basis": "《建筑抗震设计规范》第4.1.3条",
                "logic": "根据钻孔各岩土名称和土层剪切波速范围表，取得剪切波速值vs。根据波速值按以下任一条件确定覆盖层底土层：1）剪切波速Vs>500m/s且其下卧各层波速均≥500m/s；2）地面5m以下剪切波速大于其上部各土层波速2.5倍，且该层及下卧层波速均≥400m/s。计算地面至该土层顶部的距离。",
                "output": "各钻孔的覆盖层厚度。",
                "next": "进入步骤3。",
                "table": {
                    "headers": ["土的类型", "岩土名称和性状", "剪切波速范围(m/s)"],
                    "rows": [
                        ["岩石", "坚硬、较硬且完整的岩石", "vs > 800"],
                        ["坚硬土或软质岩石", "中风化砂岩、密实卵石、密实碎石土", "800 ≥ vs > 500"],
                        ["中硬土", "强风化砂岩、中密~稍密卵石、硬塑粘土", "500 ≥ vs > 250"],
                        ["中软土", "稍密砂、可塑~软塑粘土、粉土", "250 ≥ vs > 150"],
                        ["软弱土", "淤泥和淤泥质土、素填土、杂填土", "vs ≤ 150"]
                    ]
                }
            },
            {
                "step_title": "步骤3（计算等效剪切波速）",
                "input": "剪切波速值vs，覆盖层厚度",
                "basis": "《建筑抗震设计规范》第4.1.5条",
                "logic": "计算土层等效剪切波速vse：vse = d₀/t，t = Σ(di/vsi)。式中：vse为土层等效剪切波速(m/s)；d₀为计算深度(m)，取覆盖层厚度和20m两者的较小值；t为剪切波在地面至计算深度之间的传播时间；di为计算深度范围内第i土层的厚度(m)；vsi为计算深度范围内第i土层的剪切波速(m/s)；n为计算深度范围内土层的分层数。",
                "output": "等效剪切波速。",
                "next": "进入步骤4。"
            },
            {
                "step_title": "步骤4（判断场地类别）",
                "input": "覆盖层厚度和等效剪切波速",
                "basis": "《建筑抗震设计规范》第4.1.6条，表4.1.6",
                "logic": "根据等效剪切波速和覆盖层厚度，查表4.1.6获取场地类别。",
                "output": "场地类别。",
                "next": "无"
            }
        ]
    })

    # ===== 步骤三：砂土液化评价 =====
    steps.append({
        "title": "（三）砂土液化评价",
        "sub_steps": [
            {
                "step_title": "步骤1：是否进行液化判别",
                "input": "提取勘察报告2.5节场地地层岩性、4.2.1场地抗震设防烈度",
                "basis": "《建筑抗震设计规范》第4.3.1节",
                "logic": "1.场地内存在砂土和粉土；2.抗震设防烈度大于6度",
                "output": "同时满足，需进行液化判别；不同时满足，无需进行液化判别，结束。",
                "next": "同时满足，进入步骤2。"
            },
            {
                "step_title": "步骤2：液化初判（地质年代）",
                "input": "确定液化判别钻孔地层信息、地质年代、场地抗震设防烈度",
                "basis": "《建筑抗震设计规范》第4.3.3节",
                "logic": "1）地质年代为第四纪晚更新世(Q3)及其以前；2）抗震设防烈度为7、8度",
                "output": "同时满足以上2条时，可不考虑液化影响，终止；不同时满足，进入步骤3。",
                "next": "进入步骤3。"
            },
            {
                "step_title": "步骤3：液化初判（黏粒含量）",
                "input": "液化判别钻孔地层信息、抗震设防烈度、颗分试验统计结果表",
                "basis": "《建筑抗震设计规范》第4.3.3节",
                "logic": "根据颗粒分析试验统计成果表，粉土的黏粒(粒径小于0.005mm的颗粒)含量百分率，当抗震设防烈度7度、8度和9度分别不小于10、13和16时，可判为不液化土。",
                "output": "满足条件时，可不考虑液化影响；不满足时，进入步骤4。",
                "next": "进入步骤4。"
            },
            {
                "step_title": "步骤4：液化初判（覆盖层与地下水位）",
                "input": "液化判别钻孔地层信息、抗震设防烈度、颗分试验结果、抗浮水位、基础埋置深度",
                "basis": "《建筑抗震设计规范》第4.3.3节",
                "logic": "浅埋天然地基的建筑，当上覆非液化土层厚度和地下水位深度符合下列条件之一时，可不考虑液化影响：du>d₀+db-2；dw>d₀+db-3；du+dw>1.5d₀+2db-4.5。其中du为上覆非液化土层厚度(m)（扣除淤泥和淤泥质土层厚度）；dw为地下水位深度(m)；db为基础埋置深度(m)（2m以内按2m）；d₀为液化土特征深度(m)。",
                "output": "满足以上3条之一时，可判为不液化土；均不符合时，需进一步进行液化判别。",
                "next": "均不符合时，进入步骤5。",
                "table": {
                    "headers": ["饱和土类别", "7度", "8度", "9度"],
                    "rows": [
                        ["粉土", "6m", "7m", "8m"],
                        ["砂土", "7m", "8m", "9m"]
                    ]
                }
            },
            {
                "step_title": "步骤5：液化判别深度",
                "input": "勘察报告拟建建筑物性质一览表",
                "basis": "《建筑抗震设计规范》第4.2.1条",
                "logic": "建筑是否为下列之一：1.单层厂房和单层房屋；2.砌体房屋；3.不超过8层且高度在24m以下的建筑物。",
                "output": "满足以上3条之一，判别深度为15m；均不满足，判别深度为20m。",
                "next": "进入步骤6。"
            },
            {
                "step_title": "步骤6：液化复判计算",
                "input": "N₀（基准值）、ds（标贯点深度）、dw（地下水位）、ρc（黏粒含量百分率）、β（调整系数）",
                "basis": "《建筑抗震设计规范》第4.3.4条",
                "logic": "根据公式计算液化判别标准贯入锤击数临界值：Ncr = N₀×β×[ln(0.6ds+1.5)-0.1dw]×√(3/ρc)。",
                "output": "Ncr——液化判别标准贯入锤击数临界值。",
                "next": "进入步骤7。",
                "table": {
                    "headers": ["设计基本地震加速度(g)", "0.10", "0.15", "0.20", "0.30", "0.40"],
                    "rows": [["N₀", "7", "10", "12", "16", "19"]]
                }
            },
            {
                "step_title": "步骤7：得出液化复判结论",
                "input": "步骤6计算出的Ncr，深度对应的击数N（编录信息表标准贯入试验列）",
                "basis": "《建筑抗震设计规范》第4.3.4条",
                "logic": "N≤Ncr，判断为液化土；N＞Ncr，判断为不液化土。",
                "output": "N≤Ncr，判断为液化土；N＞Ncr，判断为不液化土。",
                "next": "N≤Ncr，进入步骤8；N＞Ncr，结束。"
            },
            {
                "step_title": "步骤8：液化指数计算",
                "input": "n（标贯试验点总数）、Ni和Ncri（实测值和临界值）、di（土层厚度）、Wi（权函数值）",
                "basis": "《建筑抗震设计规范》第4.3.5条",
                "logic": "根据公式计算液化指数：IlE = Σ[1-Ni/Ncri]×di×Wi。其中Wi的取值：当该层中点深度不大于5m时取10，等于20m时取零值，5~20m时按线性内插法取值。",
                "output": "IlE——液化指数。",
                "next": "进入步骤9。"
            },
            {
                "step_title": "步骤9：确定液化等级",
                "input": "液化指数IlE",
                "basis": "《建筑抗震设计规范》第4.3.5条",
                "logic": "根据液化指数确定液化等级。",
                "output": "轻微、中等或严重。",
                "next": "进入步骤10。",
                "table": {
                    "headers": ["液化等级", "轻微", "中等", "严重"],
                    "rows": [["液化指数IlE", "0 < IlE ≤ 6", "6 < IlE ≤ 18", "IlE > 18"]]
                }
            },
            {
                "step_title": "步骤10：处理措施",
                "input": "勘察报告4.2.5节抗震设防类别；步骤9计算的液化指数",
                "basis": "《建筑抗震设计规范》第4.3.6条",
                "logic": "根据表4.3.6，按抗震设防类别和液化等级采用对应的处理措施。",
                "output": "根据表4.3.6，输出对应的处理措施。",
                "next": "无",
                "table": {
                    "headers": ["设防类别", "轻微液化", "中等液化", "严重液化"],
                    "rows": [
                        ["甲类", "全部消除液化沉陷", "全部消除液化沉陷", "全部消除液化沉陷"],
                        ["乙类", "部分消除液化沉陷，或对基础和上部结构进行处理", "全部消除液化沉陷，或部分消除且对基础和上部结构进行处理", "全部消除液化沉陷"],
                        ["丙类", "可不采取措施", "对基础和上部结构进行处理，或采用更经济的措施", "全部消除液化沉陷，或部分消除且对基础和上部结构进行处理"],
                        ["丁类", "可不采取措施", "可不采取措施", "对基础和上部结构进行处理，或采用其他经济的措施"]
                    ]
                }
            }
        ]
    })

    # ===== 步骤四：软弱土震陷评价 =====
    steps.append({
        "title": "（四）软弱土震陷评价",
        "sub_steps": [
            {
                "step_title": "步骤1：是否进行震陷评价",
                "input": "勘察报告4.2.5节抗震设防类别；勘察报告2.5节",
                "basis": "《软土地区岩土工程勘察规程》第6.3.4条",
                "logic": "抗震设防烈度≥7度；存在软弱土（淤泥、淤泥质土、冲填土、杂填土及其他高压缩性土层）。",
                "output": "以上条件均满足时，进行震陷评价，进入步骤2；否则不进行。",
                "next": "均满足，进入步骤2。"
            },
            {
                "step_title": "步骤2：初判",
                "input": "勘察报告4.2.2节等效剪切波速；4.2.1节抗震设防烈度；4.4承载力特征值；3.1节塑性指数",
                "basis": "《岩土工程勘察规范(2009版)》或《软土地区岩土工程勘察规程》第6.3.4条",
                "logic": "1）当地基等效剪切波速或承载力特征值大于表中数值时，可不考虑震陷影响。2）对于饱和粉质粘土，8度(0.30g)和9度时，当塑性指数Ip<15且Ws≥0.9WL、IL≥0.75时，可判为震陷性软土。",
                "output": "地基土为震陷性软土或不考虑震陷。",
                "next": "进入步骤3。",
                "table": {
                    "headers": ["设防烈度", "等效剪切波速阈值(m/s)", "承载力特征值阈值(kPa)"],
                    "rows": [
                        ["7度(0.10g)", "90", "80"],
                        ["7度(0.15g)", "90", "80"],
                        ["8度(0.20g)", "140", "100"],
                        ["8度(0.30g)", "140", "100"],
                        ["9度(0.40g)", "200", "120"]
                    ]
                }
            },
            {
                "step_title": "步骤3：详判",
                "input": "勘察报告4.2.1节抗震设防烈度；钻孔地层信息；4.2.2节等效剪切波速",
                "basis": "《岩土工程勘察规范(2009版)》或《软土地区岩土工程勘察规程》第6.3.4条",
                "logic": "按下表取震陷估算值。",
                "output": "震陷估算值。",
                "next": "进入步骤4。",
                "table": {
                    "headers": ["设防烈度", "Vse=90~140 m/s", "Vse=140~200 m/s"],
                    "rows": [
                        ["7度(0.10g)", "30~50mm", "不考虑"],
                        ["8度(0.20g)", "50~150mm", "30~50mm"],
                        ["8度(0.30g)", "100~200mm", "50~100mm"],
                        ["9度(0.40g)", "200~400mm", "100~200mm"]
                    ]
                }
            },
            {
                "step_title": "步骤4：危害性评价及处理措施",
                "input": "震陷估算值；勘察报告第1.1.2节地基变形允许值",
                "basis": "《岩土工程勘察规范(2009版)》或《软土地区岩土工程勘察规程》第6.3.4条",
                "logic": "震陷量>50mm或地基变形允许值时，需采取措施。",
                "output": "满足条件时采取措施：采用桩基穿越软土层、深层搅拌法、高压旋喷法、换填垫层、复合地基方法加强上部结构整体性，减小不均匀沉降影响。不满足则不采取措施。",
                "next": "无"
            }
        ]
    })

    # ===== 步骤五：地震滑坡/崩塌评价 =====
    steps.append({
        "title": "（五）地震滑坡/崩塌评价",
        "sub_steps": [{
            "step_title": "地震滑坡/崩塌评价",
            "input": "勘察报告2.6节不良地质作用",
            "basis": "《建筑抗震设计规范》",
            "logic": "存在滑坡/崩塌。",
            "output": "若存在，则建议开展专项勘察或提出严禁建设或避让要求；若不存在，则说明'地震作用下不会产生滑坡、崩塌等不良地质作用'。",
            "next": "无"
        }]
    })

    # ===== 步骤六：场地横向扩展评价 =====
    steps.append({
        "title": "（六）场地横向扩展评价",
        "sub_steps": [
            {
                "step_title": "步骤1：有无液化土",
                "input": "勘察报告4.2.3地震液化评价",
                "basis": "《建筑抗震设计规范》4.3.10",
                "logic": "是否存在液化土。",
                "output": "若存在，进入步骤2；若不存在，则不考虑地震的横向扩展作用。",
                "next": "存在，进入步骤2。"
            },
            {
                "step_title": "步骤2：临空条件",
                "input": "勘察报告第2章场地环境与工程地质条件",
                "basis": "《建筑抗震设计规范》4.3.10",
                "logic": "是否满足以下临空条件之一：场地是否位于故河道，临近河岸、湖泊、海岸，边坡坡脚、坡顶和明显的地面高差、人工临空面（如深基坑）等地段。",
                "output": "若不满足，则不考虑地震的横向扩展作用；若满足，则应考虑横向扩展作用，并提出地基处理措施（振冲碎石桩、CFG、高压旋喷桩、大直径素混凝土桩、强夯、挤密砂桩，设置地下连续墙或板桩墙）、结构措施（抗滑动验算、筏板基础、地梁体系、桩基水平承载力验算）和监测方案（设置水平位移监测点，开展长期监测）。",
                "next": "无"
            }
        ]
    })

    # ===== 步骤七：抗震地段划分 =====
    steps.append({
        "title": "（七）抗震地段划分",
        "sub_steps": [
            {
                "step_title": "步骤1：危险地段判定",
                "input": "勘察报告2.8节不良地质作用",
                "basis": "《建筑抗震设计规范》4.1.1条",
                "logic": "根据不良地质作用，若地震时可能发生滑坡、崩塌、地陷、地裂、泥石流，则划分为危险地段；若不存在，则进入步骤2。",
                "output": "若存在不良地质作用，划分为危险地段；若不存在，进入步骤2。",
                "next": "不存在，进入步骤2。"
            },
            {
                "step_title": "步骤2：发震断裂判定",
                "input": "勘察报告2.3.1区域地质构造特征",
                "basis": "《建筑抗震设计规范》4.1.1条",
                "logic": "是否存在发震断裂。",
                "output": "不存在，进入步骤4；存在，进入步骤3。",
                "next": "存在，进入步骤3。"
            },
            {
                "step_title": "步骤3：发震断裂影响评估",
                "input": "勘察报告2.3.1区域地质构造特征；4.2.1抗震设防烈度",
                "basis": "《建筑抗震设计规范》4.1.1条",
                "logic": "1)抗震设防烈度小于8度；2)非全新世活动断裂；3)抗震设防烈度为8度和9度时，隐伏断裂的土层覆盖厚度分别大于60m和90m。",
                "output": "若满足任意一条，则不考虑发震断裂错动影响，进入步骤4；若均不满足，划为危险地段。",
                "next": "满足任一，进入步骤4。"
            },
            {
                "step_title": "步骤4：不利地段判定",
                "input": "勘察报告第2章场地环境与工程地质条件",
                "basis": "《建筑抗震设计规范》4.1.1条",
                "logic": "是否满足以下条件之一：①软弱土与液化土；②条状突出山嘴；③高耸孤立山丘；④非岩质的陡坡、陡坎；⑤河岸和边坡的边缘；⑥土层明显不均匀（含故河道、断层破碎带、暗埋塘浜沟谷、半填半挖地基）；⑦高含水量可塑黄土；⑧地表结构性裂缝。",
                "output": "若满足任意一条，划分为不利地段；若均不满足，进入步骤5。",
                "next": "均不满足，进入步骤5。"
            },
            {
                "step_title": "步骤5：有利地段判定",
                "input": "勘察报告第2章场地环境与工程地质条件；4.5地基与基础评价",
                "basis": "《建筑抗震设计规范》4.1.1条",
                "logic": "是否满足以下条件：①场地地基为稳定基岩、坚硬土；②地形开阔、平坦、密实、均匀的中硬土等。",
                "output": "若满足任意一条，划分为有利地段；若均不满足，划分为一般地段。",
                "next": "无"
            }
        ]
    })

    return steps


# ============================================================
# 主程序
# ============================================================
if __name__ == "__main__":
    print("=" * 60)
    print("  场地和地基地震效应评价规范 - 自动化生成脚本")
    print("=" * 60)

    # 步骤1：数据合并
    print("\n[步骤1] 合并全国抗震设防数据...")
    data = merge_data(INPUT_JSON, INPUT_TXT, OUTPUT_JSON)

    # 步骤2：生成Word文档
    print("\n[步骤2] 生成Word报告...")
    generate_docx(data, OUTPUT_DOCX)

    print("\n" + "=" * 60)
    print("  执行成功！")
    print(f"  完整JSON: {OUTPUT_JSON}")
    print(f"  Word报告: {OUTPUT_DOCX}")
    print("=" * 60)
