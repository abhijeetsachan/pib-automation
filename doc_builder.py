"""Document Generation Engine: Creates formatted Excel (.xlsx) and Word (.docx) files."""

import os
from pathlib import Path
from typing import List, Dict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from config import UPSC_PAPERS


class DocumentBuilder:
    @staticmethod
    def generate_excel(releases: List[Dict], output_filepath: str, formatted_date: str):
        """Generates a professional multi-sheet Excel file."""
        wb = openpyxl.Workbook()

        # Styles
        navy_header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        border_thin = Border(
            left=Side(style="thin", color="E2E8F0"),
            right=Side(style="thin", color="E2E8F0"),
            top=Side(style="thin", color="E2E8F0"),
            bottom=Side(style="thin", color="E2E8F0"),
        )

        # ---------------- Sheet 1: UPSC Relevant News ----------------
        ws_relevant = wb.active
        ws_relevant.title = "UPSC Relevant News"

        headers = ["S.No.", "UPSC Paper", "Ministry / Department", "Headline", "Key Topics", "Why Relevant?", "PIB Link"]
        ws_relevant.append(headers)

        for col_idx, _ in enumerate(headers, 1):
            cell = ws_relevant.cell(row=1, column=col_idx)
            cell.fill = navy_header_fill
            cell.font = header_font
            cell.alignment = align_center

        relevant_items = [r for r in releases if r.get("is_relevant")]
        row_num = 2

        for idx, item in enumerate(relevant_items, 1):
            tags_str = ", ".join(item.get("tags", []))
            paper = item.get("paper", "GS")
            ws_relevant.append([
                idx,
                paper,
                item.get("ministry", ""),
                item.get("headline", ""),
                tags_str,
                item.get("relevance_reason", ""),
                item.get("url", "")
            ])

            # Apply cell styles
            paper_color = UPSC_PAPERS.get(paper, {}).get("badge_color", "E2E8F0")
            for col_idx in range(1, 8):
                cell = ws_relevant.cell(row=row_num, column=col_idx)
                cell.border = border_thin
                cell.alignment = align_center if col_idx in (1, 2) else align_left
                if col_idx == 2:
                    cell.fill = PatternFill(start_color=paper_color, end_color=paper_color, fill_type="solid")
                    cell.font = Font(name="Calibri", bold=True)
                if col_idx == 7 and item.get("url"):
                    cell.font = Font(name="Calibri", color="2563EB", underline="single")

            row_num += 1

        # Column widths
        widths = {1: 8, 2: 15, 3: 30, 4: 55, 5: 25, 6: 35, 7: 35}
        for col, width in widths.items():
            ws_relevant.column_dimensions[get_column_letter(col)].width = width

        # ---------------- Sheet 2: All PIB Releases (Full Log) ----------------
        ws_all = wb.create_sheet(title="All PIB Releases (Full Log)")
        all_headers = ["S.No.", "Classification", "Paper", "Ministry", "Headline", "PIB URL"]
        ws_all.append(all_headers)

        for col_idx, _ in enumerate(all_headers, 1):
            cell = ws_all.cell(row=1, column=col_idx)
            cell.fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
            cell.font = header_font
            cell.alignment = align_center

        for idx, item in enumerate(releases, 1):
            status = "Relevant (UPSC)" if item.get("is_relevant") else "Routine Fluff"
            ws_all.append([
                idx,
                status,
                item.get("paper", "N/A"),
                item.get("ministry", ""),
                item.get("headline", ""),
                item.get("url", "")
            ])

        for col, width in {1: 8, 2: 18, 3: 15, 4: 30, 5: 60, 6: 35}.items():
            ws_all.column_dimensions[get_column_letter(col)].width = width

        wb.save(output_filepath)
        return output_filepath

    @staticmethod
    def generate_docx(releases: List[Dict], output_filepath: str, formatted_date: str):
        """Generates an executive Word document tailored for UPSC civil services preparation."""
        doc = Document()

        # Margins
        for section in doc.sections:
            section.top_margin = Inches(0.8)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(0.8)
            section.right_margin = Inches(0.8)

        # Title
        p_title = doc.add_paragraph()
        p_title.paragraph_format.space_before = Pt(0)
        p_title.paragraph_format.space_after = Pt(2)
        r_title = p_title.add_run("PIB DAILY UPSC COMPILATION")
        r_title.font.name = "Arial"
        r_title.font.size = Pt(20)
        r_title.font.bold = True
        r_title.font.color.rgb = RGBColor(30, 58, 138)  # Deep Navy

        # Subtitle
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_after = Pt(14)
        r_sub = p_sub.add_run(f"Date: {formatted_date}  |  Target: UPSC CSE (Prelims & Mains)")
        r_sub.font.name = "Arial"
        r_sub.font.size = Pt(10.5)
        r_sub.font.italic = True
        r_sub.font.color.rgb = RGBColor(71, 85, 105)

        # Executive Metrics Box
        relevant_items = [r for r in releases if r.get("is_relevant")]
        summary_table = doc.add_table(rows=1, cols=1)
        summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = summary_table.cell(0, 0)
        
        counts_by_paper = {}
        for r in relevant_items:
            p = r.get("paper", "Other")
            counts_by_paper[p] = counts_by_paper.get(p, 0) + 1
        dist_str = ", ".join([f"{k}: {v}" for k, v in sorted(counts_by_paper.items())]) or "None"

        summary_p = cell.paragraphs[0]
        summary_p.paragraph_format.space_before = Pt(4)
        summary_p.paragraph_format.space_after = Pt(4)
        r_sum = summary_p.add_run(
            f"DAILY SUMMARY: Scraped {len(releases)} total PIB releases today. "
            f"Found {len(relevant_items)} high-yield articles aligned with the UPSC syllabus ({dist_str})."
        )
        r_sum.font.size = Pt(10)
        r_sum.font.bold = True

        doc.add_paragraph().paragraph_format.space_after = Pt(8)

        # Group by GS Paper
        ordered_papers = ["GS-1", "GS-2", "GS-3", "GS-4"]
        grouped = {p: [] for p in ordered_papers}
        for item in relevant_items:
            p = item.get("paper", "")
            if p in grouped:
                grouped[p].append(item)

        for paper_code in ordered_papers:
            items = grouped[paper_code]
            if not items:
                continue

            paper_meta = UPSC_PAPERS.get(paper_code, {})
            h_p = doc.add_paragraph()
            h_p.paragraph_format.space_before = Pt(16)
            h_p.paragraph_format.space_after = Pt(4)

            r_h = h_p.add_run(f"■ {paper_meta.get('title', paper_code)}: {paper_meta.get('subtopics', '')}")
            r_h.font.name = "Arial"
            r_h.font.size = Pt(13)
            r_h.font.bold = True
            r_h.font.color.rgb = RGBColor(37, 99, 235)  # Royal Blue

            for idx, item in enumerate(items, 1):
                p_item = doc.add_paragraph()
                p_item.paragraph_format.left_indent = Inches(0.2)
                p_item.paragraph_format.space_before = Pt(4)
                p_item.paragraph_format.space_after = Pt(2)

                # Number & Headline
                r_num = p_item.add_run(f"{idx}. ")
                r_num.font.bold = True
                r_hl = p_item.add_run(f"{item['headline']}\n")
                r_hl.font.name = "Calibri"
                r_hl.font.size = Pt(11.5)
                r_hl.font.bold = True

                # Ministry & Tags
                p_meta = doc.add_paragraph()
                p_meta.paragraph_format.left_indent = Inches(0.4)
                p_meta.paragraph_format.space_after = Pt(2)

                r_min = p_meta.add_run(f"Ministry: {item['ministry']}  |  Keywords: {', '.join(item.get('tags', []))}\n")
                r_min.font.size = Pt(9.5)
                r_min.font.italic = True
                r_min.font.color.rgb = RGBColor(100, 116, 139)

                # Syllabus Alignment Reason
                r_rel = p_meta.add_run(f"UPSC Alignment: {item.get('relevance_reason', '')}\n")
                r_rel.font.size = Pt(10)
                r_rel.font.bold = True
                r_rel.font.color.rgb = RGBColor(30, 64, 175)

                # Link
                r_url = p_meta.add_run(f"Official PIB Link: {item['url']}\n")
                r_url.font.size = Pt(9)
                r_url.font.color.rgb = RGBColor(2, 132, 199)

        doc.save(output_filepath)
        return output_filepath
