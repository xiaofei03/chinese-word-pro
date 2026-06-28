from __future__ import annotations

import argparse
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


def set_run_fonts(run, east_asia: str, latin: str, size_pt: float | None = None, bold=None, italic=None, subscript=False):
    run.font.name = latin
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), east_asia)
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    run.font.subscript = subscript


def set_style_fonts(doc: Document, lang: str):
    if lang == "cn":
        mappings = {
            "Normal": ("宋体", "Times New Roman", 10.5, False),
            "Body Text": ("宋体", "Times New Roman", 10.5, False),
            "First Paragraph": ("宋体", "Times New Roman", 10.5, False),
            "Heading 1": ("黑体", "Times New Roman", 15, True),
            "Heading 2": ("黑体", "Times New Roman", 13, True),
            "Heading 3": ("黑体", "Times New Roman", 12, True),
            "Caption": ("宋体", "Times New Roman", 10.5, False),
        }
    else:
        mappings = {
            "Normal": ("Times New Roman", "Times New Roman", 11, False),
            "Body Text": ("Times New Roman", "Times New Roman", 11, False),
            "First Paragraph": ("Times New Roman", "Times New Roman", 11, False),
            "Heading 1": ("Times New Roman", "Times New Roman", 15, True),
            "Heading 2": ("Times New Roman", "Times New Roman", 13, True),
            "Heading 3": ("Times New Roman", "Times New Roman", 12, True),
            "Caption": ("Times New Roman", "Times New Roman", 10.5, False),
        }

    for style_name, (east_asia, latin, size_pt, bold) in mappings.items():
        try:
            style = doc.styles[style_name]
        except KeyError:
            continue
        style.font.name = latin
        style._element.rPr.rFonts.set(qn("w:eastAsia"), east_asia)
        style.font.size = Pt(size_pt)
        style.font.bold = bold


def paragraph_has_drawing(paragraph) -> bool:
    return bool(paragraph._element.xpath(".//w:drawing"))


def clear_paragraph(paragraph):
    p = paragraph._element
    for child in list(p):
        p.remove(child)


def append_run(paragraph, text: str, east_asia: str, latin: str, size_pt: float, *, bold=None, italic=None, subscript=False):
    run = paragraph.add_run(text)
    set_run_fonts(run, east_asia, latin, size_pt, bold=bold, italic=italic, subscript=subscript)
    return run


def append_symbol(paragraph, base: str, sub: str | None, east_asia: str, latin: str, size_pt: float, *, italic=True, greek=False):
    append_run(paragraph, base, east_asia, latin, size_pt, italic=(False if greek else italic))
    if sub:
        append_run(paragraph, sub, east_asia, latin, size_pt - 1, italic=(False if greek else italic), subscript=True)


def _m_el(tag: str):
    return OxmlElement(f"m:{tag}")


def _w_el(tag: str):
    return OxmlElement(f"w:{tag}")


def math_text(text: str):
    run = _m_el("r")
    r_pr = _m_el("rPr")
    r_pr.append(_m_el("nor"))
    run.append(r_pr)
    t = _m_el("t")
    t.text = text
    run.append(t)
    return run


def math_rich_text(text: str, *, italic: bool = False):
    run = _m_el("r")
    r_pr = _m_el("rPr")
    if not italic:
        r_pr.append(_m_el("nor"))
    run.append(r_pr)
    t = _m_el("t")
    t.text = text
    run.append(t)
    return run


def math_sub(base: str, sub: str, *, italic_base: bool = False):
    node = _m_el("sSub")
    e = _m_el("e")
    e.append(math_rich_text(base, italic=italic_base))
    sub_el = _m_el("sub")
    sub_el.append(math_rich_text(sub, italic=False))
    node.append(e)
    node.append(sub_el)
    return node


def math_subsup(base: str, sub: str, sup: str, *, italic_base: bool = False):
    node = _m_el("sSubSup")
    e = _m_el("e")
    e.append(math_rich_text(base, italic=italic_base))
    sub_el = _m_el("sub")
    sub_el.append(math_rich_text(sub, italic=False))
    sup_el = _m_el("sup")
    sup_el.append(math_rich_text(sup, italic=False))
    node.append(e)
    node.append(sub_el)
    node.append(sup_el)
    return node


def math_group(elements):
    group = _m_el("d")
    d_pr = _m_el("dPr")
    beg = _m_el("begChr")
    beg.set(qn("m:val"), "(")
    end = _m_el("endChr")
    end.set(qn("m:val"), ")")
    d_pr.append(beg)
    d_pr.append(end)
    group.append(d_pr)
    e = _m_el("e")
    for element in elements:
        e.append(element)
    group.append(e)
    return group


def build_omml_multiline_paragraph(lines):
    o_math_para = _m_el("oMathPara")
    o_math_para_pr = _m_el("oMathParaPr")
    jc_math = _m_el("jc")
    jc_math.set(qn("m:val"), "center")
    o_math_para_pr.append(jc_math)
    o_math_para.append(o_math_para_pr)
    o_math = _m_el("oMath")
    eq_arr = _m_el("eqArr")
    eq_arr_pr = _m_el("eqArrPr")
    for tag, value in (("maxDist", "1"), ("objDist", "0"), ("rSp", "1")):
        node = _m_el(tag)
        node.set(qn("m:val"), value)
        eq_arr_pr.append(node)
    eq_arr.append(eq_arr_pr)
    for line in lines:
        e = _m_el("e")
        for comp in line:
            e.append(comp)
        eq_arr.append(e)
    o_math.append(eq_arr)
    o_math_para.append(o_math)
    p_pr = _w_el("pPr")
    jc = _w_el("jc")
    jc.set(qn("w:val"), "center")
    p_pr.append(jc)
    para = OxmlElement("w:p")
    para.append(p_pr)
    para.append(o_math_para)
    return para


def build_text_paragraph(text: str, align: str = "right"):
    para = OxmlElement("w:p")
    p_pr = _w_el("pPr")
    jc = _w_el("jc")
    jc.set(qn("w:val"), align)
    p_pr.append(jc)
    para.append(p_pr)
    run = _w_el("r")
    r_pr = _w_el("rPr")
    r_fonts = _w_el("rFonts")
    r_fonts.set(qn("w:ascii"), "Times New Roman")
    r_fonts.set(qn("w:hAnsi"), "Times New Roman")
    r_fonts.set(qn("w:eastAsia"), "宋体")
    r_pr.append(r_fonts)
    sz = _w_el("sz")
    sz.set(qn("w:val"), "21")
    r_pr.append(sz)
    run.append(r_pr)
    t = _w_el("t")
    t.text = text
    run.append(t)
    para.append(run)
    return para


def replace_with_omml_block(paragraph, lines, eq_no: int | None = None):
    parent = paragraph._element.getparent()
    new_p = build_omml_multiline_paragraph(lines)
    paragraph._element.addnext(new_p)
    if eq_no is not None:
        new_p.addnext(build_text_paragraph(f"（{eq_no}）", align="right"))
    parent.remove(paragraph._element)


def set_paragraph_format(paragraph, lang: str, in_table: bool = False):
    pf = paragraph.paragraph_format
    pf.left_indent = Pt(0)
    pf.right_indent = Pt(0)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if in_table:
        pf.first_line_indent = Pt(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(16 if lang == "cn" else 14)
        return

    text = paragraph.text.strip()
    style_name = paragraph.style.name if paragraph.style else ""
    has_drawing = paragraph_has_drawing(paragraph)
    if has_drawing:
        pf.first_line_indent = Pt(0)
        pf.space_before = Pt(6)
        pf.space_after = Pt(6)
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.line_spacing = None
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return

    if style_name.startswith("Heading"):
        pf.first_line_indent = Pt(0)
        pf.space_before = Pt(12)
        pf.space_after = Pt(6)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(20 if lang == "cn" else 16)
    elif text.startswith(("表 ", "图 ", "Table ", "Figure ")):
        pf.first_line_indent = Pt(0)
        pf.space_before = Pt(6)
        pf.space_after = Pt(6)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(18 if lang == "cn" else 14)
    else:
        pf.first_line_indent = Cm(0.74) if lang == "cn" else Pt(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY if lang == "cn" else WD_LINE_SPACING.ONE_POINT_FIVE
        pf.line_spacing = Pt(20 if lang == "cn" else 18)


def set_table_borders(cell, *, top=None, bottom=None, left=None, right=None):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)

    def apply(edge: str, spec):
        node = borders.find(qn(f"w:{edge}"))
        if spec is None:
            if node is not None:
                borders.remove(node)
            return
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        for k, v in spec.items():
            node.set(qn(f"w:{k}"), str(v))

    apply("top", top)
    apply("bottom", bottom)
    apply("left", left)
    apply("right", right)


def format_tables(doc: Document, lang: str):
    top_rule = {"val": "single", "sz": "12", "space": "0", "color": "000000"}
    mid_rule = {"val": "single", "sz": "4", "space": "0", "color": "000000"}
    bottom_rule = {"val": "single", "sz": "12", "space": "0", "color": "000000"}
    none_rule = {"val": "nil"}
    east_asia = "宋体" if lang == "cn" else "Times New Roman"
    latin = "Times New Roman"
    body_size = 10.5 if lang == "cn" else 10.5

    for table in doc.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True
        last_row = len(table.rows) - 1
        for r_idx, row in enumerate(table.rows):
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for p in cell.paragraphs:
                    set_paragraph_format(p, lang, in_table=True)
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    for run in p.runs:
                        set_run_fonts(run, east_asia, latin, body_size, bold=(r_idx == 0))
                set_table_borders(
                    cell,
                    top=top_rule if r_idx == 0 else none_rule,
                    bottom=mid_rule if r_idx == 0 else (bottom_rule if r_idx == last_row else none_rule),
                    left=none_rule,
                    right=none_rule,
                )


def rebuild_equation_paragraph(paragraph):
    text = paragraph.text.strip()
    if text.startswith("W_kl = sum_p I(k in p) I(l in p), k != l"):
        replace_with_omml_block(
            paragraph,
            [[
                math_sub("W", "kl"),
                math_text(" = "),
                math_sub("Σ", "p"),
                math_text(" I"),
                math_group([math_text("k in p")]),
                math_text(" I"),
                math_group([math_text("l in p")]),
                math_text(", k ≠ l"),
            ]],
            eq_no=1,
        )
        return True
    if text.startswith("PR_k = (1-d)/N + d sum_l W_lk PR_l / sum_m W_lm"):
        replace_with_omml_block(
            paragraph,
            [
                [
                    math_sub("PR", "k"),
                    math_text(" = "),
                    math_group([math_text("1-d")]),
                    math_text("/N + d "),
                    math_sub("Σ", "l"),
                    math_text(" "),
                    math_sub("W", "lk"),
                    math_text(" "),
                    math_sub("PR", "l"),
                    math_text(" / "),
                    math_sub("Σ", "m"),
                    math_text(" "),
                    math_sub("W", "lm"),
                ],
            ],
            eq_no=2,
        )
        return True
    if text.startswith("AIPatent_it = sum_k N_ikt PR_kt"):
        replace_with_omml_block(
            paragraph,
            [[
                math_sub("AIPatent", "it"),
                math_text(" = "),
                math_sub("Σ", "k"),
                math_text(" "),
                math_sub("N", "ikt"),
                math_text(" "),
                math_sub("PR", "kt"),
            ]],
            eq_no=3,
        )
        return True
    if text.startswith("AIW_it = z(AIDisclosure_it) - z(AIPatent_it)"):
        replace_with_omml_block(
            paragraph,
            [[
                math_sub("AIW", "it"),
                math_text(" = z"),
                math_group([math_sub("AIDisclosure", "it")]),
                math_text(" - z"),
                math_group([math_sub("AIPatent", "it")]),
            ]],
            eq_no=4,
        )
        return True
    if text.startswith("Resilience_it = alpha_0 + alpha_1 AIW_it"):
        replace_with_omml_block(
            paragraph,
            [
                [
                    math_sub("Resilience", "it"),
                    math_text(" = "),
                    math_sub("α", "0"),
                    math_text(" + "),
                    math_sub("α", "1"),
                    math_text(" "),
                    math_sub("AIW", "it"),
                    math_text(" + "),
                    math_sub("α", "2"),
                    math_text(" "),
                    math_subsup("AIW", "it", "2"),
                    math_text(" + Σ "),
                    math_sub("β", "k"),
                    math_text(" "),
                    math_sub("Controls", "it"),
                ],
                [
                    math_text("+ "),
                    math_sub("μ", "i"),
                    math_text(" + "),
                    math_sub("λ", "t"),
                    math_text(" + "),
                    math_sub("ε", "it"),
                ],
            ],
            eq_no=5,
        )
        return True
    if text.startswith("Channel_it = gamma_0 + gamma_1 AIW_it"):
        replace_with_omml_block(
            paragraph,
            [
                [
                    math_sub("Channel", "it"),
                    math_text(" = "),
                    math_sub("γ", "0"),
                    math_text(" + "),
                    math_sub("γ", "1"),
                    math_text(" "),
                    math_sub("AIW", "it"),
                    math_text(" + "),
                    math_sub("γ", "2"),
                    math_text(" "),
                    math_subsup("AIW", "it", "2"),
                    math_text(" + Σ "),
                    math_sub("δ", "k"),
                    math_text(" "),
                    math_sub("Controls", "it"),
                ],
                [
                    math_text("+ "),
                    math_sub("μ", "i"),
                    math_text(" + "),
                    math_sub("λ", "t"),
                    math_text(" + "),
                    math_sub("ν", "it"),
                ],
            ],
            eq_no=6,
        )
        return True
    if text.startswith("Resilience_it = theta_0 + theta_1 AIW_it"):
        replace_with_omml_block(
            paragraph,
            [
                [
                    math_sub("Resilience", "it"),
                    math_text(" = "),
                    math_sub("θ", "0"),
                    math_text(" + "),
                    math_sub("θ", "1"),
                    math_text(" "),
                    math_sub("AIW", "it"),
                    math_text(" + "),
                    math_sub("θ", "2"),
                    math_text(" "),
                    math_subsup("AIW", "it", "2"),
                    math_text(" + "),
                    math_sub("θ", "3"),
                    math_text(" "),
                    math_sub("Channel", "it"),
                ],
                [
                    math_text("+ Σ "),
                    math_sub("φ", "k"),
                    math_text(" "),
                    math_sub("Controls", "it"),
                    math_text(" + "),
                    math_sub("μ", "i"),
                    math_text(" + "),
                    math_sub("λ", "t"),
                    math_text(" + "),
                    math_sub("ξ", "it"),
                ],
            ],
            eq_no=7,
        )
        return True
    if text.startswith("Outcome_it = rho_0 + rho_1 AIW_it"):
        replace_with_omml_block(
            paragraph,
            [
                [
                    math_sub("Outcome", "it"),
                    math_text(" = "),
                    math_sub("ρ", "0"),
                    math_text(" + "),
                    math_sub("ρ", "1"),
                    math_text(" "),
                    math_sub("AIW", "it"),
                    math_text(" + "),
                    math_sub("ρ", "2"),
                    math_text(" "),
                    math_subsup("AIW", "it", "2"),
                    math_text(" + "),
                    math_sub("ρ", "3"),
                    math_text(" "),
                    math_sub("AnalystAttention", "it"),
                ],
                [
                    math_text("+ "),
                    math_sub("ρ", "4"),
                    math_text(" "),
                    math_group([math_sub("AIW", "it"), math_text(" × "), math_sub("AnalystAttention", "it")]),
                    math_text(" + "),
                    math_sub("ρ", "5"),
                    math_text(" "),
                    math_group([math_subsup("AIW", "it", "2"), math_text(" × "), math_sub("AnalystAttention", "it")]),
                ],
                [
                    math_text("+ Σ "),
                    math_sub("κ", "k"),
                    math_text(" "),
                    math_sub("Controls", "it"),
                    math_text(" + "),
                    math_sub("μ", "i"),
                    math_text(" + "),
                    math_sub("λ", "t"),
                    math_text(" + "),
                    math_sub("ω", "it"),
                ],
            ],
            eq_no=8,
        )
        return True
    return False


def rebuild_explanation_paragraph(paragraph, lang: str):
    east_asia = "宋体" if lang == "cn" else "Times New Roman"
    latin = "Times New Roman"
    size_pt = 10.5 if lang == "cn" else 11
    text = paragraph.text.strip()
    if lang == "cn":
        if text.startswith("其中，`W_kl` 表示节点"):
            clear_paragraph(paragraph)
            append_run(paragraph, "其中，", east_asia, latin, size_pt)
            append_symbol(paragraph, "W", "kl", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " 表示节点 k 与节点 l 的共现强度，I(k in p) 表示专利 p 是否包含技术节点 k。这一矩阵并不只记录专利数量，而是刻画 AI 技术知识在全球专利中的组合关系：两个 IPC 节点共同出现在 AI 专利中的频次越高，说明相应技术方向之间的知识关联越紧密。", east_asia, latin, size_pt)
            return True
        if text.startswith("其中，`PR_k` 表示技术节点"):
            clear_paragraph(paragraph)
            append_run(paragraph, "其中，", east_asia, latin, size_pt)
            append_symbol(paragraph, "PR", "k", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " 表示技术节点 k 的 PageRank 权重，N 为网络中的技术节点数量，d 为阻尼系数，", east_asia, latin, size_pt)
            append_symbol(paragraph, "W", "lk", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " 表示节点 l 指向节点 k 的共现权重，", east_asia, latin, size_pt)
            append_symbol(paragraph, "Σ", "m", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " ", east_asia, latin, size_pt)
            append_symbol(paragraph, "W", "lm", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " 表示节点 l 的加权外向连接总强度。为反映技术网络结构随时间演化的特征，本文优先采用当年全球 AI 专利网络计算得到的节点权重；当某些年份节点信息不足时，采用全样本期网络权重进行补充。由此得到的权重可以区分不同 AI 专利所处技术位置的质量差异。", east_asia, latin, size_pt)
            return True
        if text.startswith("其中，`PR_kt` 为技术节点"):
            clear_paragraph(paragraph)
            append_run(paragraph, "其中，", east_asia, latin, size_pt)
            append_symbol(paragraph, "PR", "kt", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " 为技术节点 k 在年份 t 的网络权重。该指标的含义是：企业并非因为拥有更多 AI 专利就自动获得更高技术行动得分，而是当其 AI 专利集中在全球 AI 知识网络中更核心、更具结构影响力的技术节点时，才获得更高的质量调整后 AI 技术积累得分。因此，AI Patent 衡量的是企业可验证 AI 技术资产的网络加权存量，而不是未经区分的专利件数。", east_asia, latin, size_pt)
            return True
        if text.startswith("其中，z(AIDisclosure_it) 和 z(AIPatent_it)"):
            clear_paragraph(paragraph)
            append_run(paragraph, "其中，z(", east_asia, latin, size_pt)
            append_symbol(paragraph, "AIDisclosure", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, ") 和 z(", east_asia, latin, size_pt)
            append_symbol(paragraph, "AIPatent", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, ") 分别表示 AI Disclosure 与 AI Patent 在年份 t 内的横截面标准化值。AIW 数值越大，表示企业人工智能叙事越明显领先于其实质性人工智能技术积累，人工智能洗白程度越高。需要说明的是，这一指标本质上刻画的是“AI 叙事 - 能力基础偏离”而非对企业主观意图的直接观测，因此更接近 AI washing 的可操作代理变量，而不是对“误导行为”本身的完全识别。之所以仍将其用于度量 AI washing，是因为在上市公司情境下，外部受众最先观察到的正是公开叙事，而专利积累虽然不能穷尽企业全部 AI 能力，却能够较稳定地反映其可验证技术行动基础。当企业 AI 叙事持续显著快于这一可验证基础时，外部受众面对的正是本文所关注的“高叙事可见性、低能力可即时验证性”的偏离状态。为识别潜在的非线性关系，本文进一步构造平方项 ", east_asia, latin, size_pt)
            append_run(paragraph, "AIW", east_asia, latin, size_pt)
            append_run(paragraph, "²", east_asia, latin, size_pt)
            append_run(paragraph, "。", east_asia, latin, size_pt)
            return True
        if text.startswith("其中，Resilience_it 表示企业 i 在年份 t 的企业韧性"):
            clear_paragraph(paragraph)
            append_run(paragraph, "其中，", east_asia, latin, size_pt)
            append_symbol(paragraph, "Resilience", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " 表示企业 i 在年份 t 的企业韧性；", east_asia, latin, size_pt)
            append_symbol(paragraph, "AIW", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " 为人工智能洗白程度；", east_asia, latin, size_pt)
            append_symbol(paragraph, "AIW", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, "² 为其平方项；", east_asia, latin, size_pt)
            append_symbol(paragraph, "Controls", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " 为控制变量集合；", east_asia, latin, size_pt)
            append_symbol(paragraph, "μ", "i", east_asia, latin, size_pt, greek=True)
            append_run(paragraph, " 为企业固定效应；", east_asia, latin, size_pt)
            append_symbol(paragraph, "λ", "t", east_asia, latin, size_pt, greek=True)
            append_run(paragraph, " 为年份固定效应；", east_asia, latin, size_pt)
            append_symbol(paragraph, "ε", "it", east_asia, latin, size_pt, greek=True)
            append_run(paragraph, " 为随机扰动项。若 ", east_asia, latin, size_pt)
            append_symbol(paragraph, "α", "1", east_asia, latin, size_pt, greek=True)
            append_run(paragraph, " > 0 且 ", east_asia, latin, size_pt)
            append_symbol(paragraph, "α", "2", east_asia, latin, size_pt, greek=True)
            append_run(paragraph, " < 0，则表明人工智能洗白与企业韧性之间存在显著的倒 U 型关系。", east_asia, latin, size_pt)
            return True
        if text.startswith("其中，Channel_it 分别表示 Trade Credit 和 Agency Cost"):
            clear_paragraph(paragraph)
            append_run(paragraph, "其中，", east_asia, latin, size_pt)
            append_symbol(paragraph, "Channel", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " 分别表示 Trade Credit 和 Agency Cost。在此基础上，再将中介变量纳入企业韧性回归：", east_asia, latin, size_pt)
            return True
        if text.startswith("其中，Outcome_it 分别表示 Resilience、Trade Credit 和 Agency Cost"):
            clear_paragraph(paragraph)
            append_run(paragraph, "其中，", east_asia, latin, size_pt)
            append_symbol(paragraph, "Outcome", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " 分别表示 Resilience、Trade Credit 和 Agency Cost。为降低交互项带来的多重共线性问题，本文在构造交互项前对 AIW 与 Analyst Attention 进行中心化处理。该模型不仅能够检验分析师关注度是否改变人工智能洗白的边际影响方向与强度，也能够识别其是否改变倒 U 型关系的拐点位置。", east_asia, latin, size_pt)
            return True
    else:
        if text.startswith("where `W_kl` is the co-occurrence strength"):
            clear_paragraph(paragraph)
            append_run(paragraph, "where ", east_asia, latin, size_pt)
            append_symbol(paragraph, "W", "kl", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " is the co-occurrence strength between nodes k and l, and I(k in p) indicates whether patent p contains technology node k. The matrix therefore captures how AI-related technological knowledge is combined across patent classes. A higher co-occurrence frequency indicates a tighter knowledge association between the corresponding technology areas.", east_asia, latin, size_pt)
            return True
        if text.startswith("where `PR_k` is the PageRank weight"):
            clear_paragraph(paragraph)
            append_run(paragraph, "where ", east_asia, latin, size_pt)
            append_symbol(paragraph, "PR", "k", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " is the PageRank weight of technology node k, N is the number of nodes in the network, d is the damping factor, ", east_asia, latin, size_pt)
            append_symbol(paragraph, "W", "lk", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " is the weighted co-occurrence link from node l to node k, and ", east_asia, latin, size_pt)
            append_symbol(paragraph, "Σ", "m", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " ", east_asia, latin, size_pt)
            append_symbol(paragraph, "W", "lm", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " is the total weighted outward connection strength of node l. To reflect the evolution of the technology network over time, we use year-specific PageRank weights whenever available and supplement them with full-sample network weights when node information is insufficient in a given year. These weights distinguish AI patents by the structural importance of the technological knowledge they embody.", east_asia, latin, size_pt)
            return True
        if text.startswith("where `PR_kt` is the network weight"):
            clear_paragraph(paragraph)
            append_run(paragraph, "where ", east_asia, latin, size_pt)
            append_symbol(paragraph, "PR", "kt", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " is the network weight of node k in year t. This score means that a firm does not receive a higher AI-action score simply because it files more AI patents. It receives a higher score when its AI patents are concentrated in more central and structurally influential positions within the global AI knowledge network. The resulting AI Patent variable therefore measures a firm's quality-adjusted stock of verifiable AI technological assets rather than an undifferentiated patent count.", east_asia, latin, size_pt)
            return True
        if text.startswith("where z(AIDisclosure_it) and z(AIPatent_it)"):
            clear_paragraph(paragraph)
            append_run(paragraph, "where z(", east_asia, latin, size_pt)
            append_symbol(paragraph, "AIDisclosure", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, ") and z(", east_asia, latin, size_pt)
            append_symbol(paragraph, "AIPatent", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, ") denote the within-year cross-sectional standardized values of AI Disclosure and AI Patent, respectively. Higher values of AIW indicate that a firm's AI-oriented narrative is more pronounced relative to its substantive AI technological accumulation. Importantly, this measure captures rhetoric-capability divergence rather than managerial intent directly. It is therefore best interpreted as an operational proxy for AI washing rather than a complete observation of deceptive behavior itself. We nonetheless use it as our empirical measure because, in listed-firm settings, outside audiences first observe public rhetoric, whereas patent accumulation, while not exhausting all AI capability, provides a relatively stable indicator of verifiable technological action. When AI rhetoric persistently and substantially outruns that verifiable base, outside audiences face precisely the high-visibility, low-immediate-verifiability condition that defines the phenomenon of interest here. To test for nonlinearity, we further construct the squared term ", east_asia, latin, size_pt)
            append_run(paragraph, "AIW", east_asia, latin, size_pt)
            append_run(paragraph, "²", east_asia, latin, size_pt)
            append_run(paragraph, ".", east_asia, latin, size_pt)
            return True
        if text.startswith("where Resilience_it denotes corporate resilience for firm i in year t"):
            clear_paragraph(paragraph)
            append_run(paragraph, "where ", east_asia, latin, size_pt)
            append_symbol(paragraph, "Resilience", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " denotes corporate resilience for firm i in year t; ", east_asia, latin, size_pt)
            append_symbol(paragraph, "AIW", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " is the degree of AI washing; ", east_asia, latin, size_pt)
            append_symbol(paragraph, "AIW", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, "² is its squared term; ", east_asia, latin, size_pt)
            append_symbol(paragraph, "Controls", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " is a vector of controls; ", east_asia, latin, size_pt)
            append_symbol(paragraph, "μ", "i", east_asia, latin, size_pt, greek=True)
            append_run(paragraph, " denotes firm fixed effects; ", east_asia, latin, size_pt)
            append_symbol(paragraph, "λ", "t", east_asia, latin, size_pt, greek=True)
            append_run(paragraph, " denotes year fixed effects; and ", east_asia, latin, size_pt)
            append_symbol(paragraph, "ε", "it", east_asia, latin, size_pt, greek=True)
            append_run(paragraph, " is the disturbance term. An inverted U-shaped relationship is supported when ", east_asia, latin, size_pt)
            append_symbol(paragraph, "α", "1", east_asia, latin, size_pt, greek=True)
            append_run(paragraph, " > 0 and ", east_asia, latin, size_pt)
            append_symbol(paragraph, "α", "2", east_asia, latin, size_pt, greek=True)
            append_run(paragraph, " < 0.", east_asia, latin, size_pt)
            return True
        if text.startswith("where Channel_it is alternately Trade Credit and Agency Cost"):
            clear_paragraph(paragraph)
            append_run(paragraph, "where ", east_asia, latin, size_pt)
            append_symbol(paragraph, "Channel", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " is alternately Trade Credit and Agency Cost. We then add the mediator to the resilience equation:", east_asia, latin, size_pt)
            return True
        if text.startswith("where Outcome_it is alternately Resilience, Trade Credit, and Agency Cost"):
            clear_paragraph(paragraph)
            append_run(paragraph, "where ", east_asia, latin, size_pt)
            append_symbol(paragraph, "Outcome", "it", east_asia, latin, size_pt, italic=False)
            append_run(paragraph, " is alternately Resilience, Trade Credit, and Agency Cost. To reduce multicollinearity in the interaction terms, we center AIW and Analyst Attention before constructing the interaction terms. This specification allows us to test not only whether analyst attention changes the marginal effect of AI washing, but also whether it shifts the turning point of the nonlinear relationship.", east_asia, latin, size_pt)
            return True
    return False


def format_paragraphs(doc: Document, lang: str):
    east_asia = "宋体" if lang == "cn" else "Times New Roman"
    latin = "Times New Roman"
    body_size = 10.5 if lang == "cn" else 11
    for paragraph in list(doc.paragraphs):
        if rebuild_equation_paragraph(paragraph):
            continue
        if rebuild_explanation_paragraph(paragraph, lang):
            set_paragraph_format(paragraph, lang, in_table=False)
            continue
        set_paragraph_format(paragraph, lang, in_table=False)
        for run in paragraph.runs:
            if paragraph_has_drawing(paragraph):
                continue
            size = body_size
            style_name = paragraph.style.name if paragraph.style else ""
            if style_name == "Heading 1":
                size = 15
            elif style_name == "Heading 2":
                size = 13
            elif style_name == "Heading 3":
                size = 12
            elif paragraph.text.strip().startswith(("表 ", "图 ", "Table ", "Figure ")):
                size = 10.5
            set_run_fonts(run, east_asia, latin, size)


def apply_page_break_rules(doc: Document, lang: str):
    if lang == "cn":
        break_before_titles = {"摘要", "引言", "理论基础与研究假设", "研究设计", "实证结果分析", "进一步分析", "讨论", "结论", "参考文献"}
    else:
        break_before_titles = {"Abstract", "Introduction", "Theory and Hypotheses", "Research Design", "Empirical Results", "Additional Analyses", "Discussion", "Conclusion", "References", "参考文献"}
    for paragraph in doc.paragraphs:
        paragraph.paragraph_format.page_break_before = paragraph.text.strip() in break_before_titles


def validate_docx(path: Path):
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
    if "�" in xml or "????" in xml:
        raise RuntimeError(f"{path} contains garbling markers")
    if "ADDIN ZOTERO_ITEM" not in xml and "CSL_CITATION" not in xml:
        raise RuntimeError(f"{path} lost live citation fields")
    for old_name in ("CR_it", "Y_it", "Mediator_it", "ROA1", "Balance1", "Mshare", "Occupy"):
        if old_name in xml:
            raise RuntimeError(f"{path} still contains deprecated variable token: {old_name}")
    if "<wp:anchor" in xml:
        raise RuntimeError(f"{path} contains floating anchor images")


def main():
    parser = argparse.ArgumentParser(description="Finalize bilingual submission DOCX with citation-safe post-processing.")
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--lang", choices=["cn", "en"], required=True)
    parser.add_argument("--mode", default="journal_submission")
    parser.add_argument("--reference-style", default=None)
    parser.add_argument("--template-profile", default=None)
    args = parser.parse_args()
    if args.mode != "journal_submission":
        raise ValueError("Only --mode journal_submission is currently supported")
    doc = Document(args.input_docx)
    set_style_fonts(doc, args.lang)
    format_paragraphs(doc, args.lang)
    apply_page_break_rules(doc, args.lang)
    format_tables(doc, args.lang)
    out_path = Path(args.output_docx)
    doc.save(out_path)
    validate_docx(out_path)


if __name__ == "__main__":
    main()
