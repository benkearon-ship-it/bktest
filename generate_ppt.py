#!/usr/bin/env python3
"""
Generate PowerPoint presentation from TD Bank Account Plan markdown
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import re

# Harness brand colors
HARNESS_PURPLE = RGBColor(107, 70, 193)  # #6B46C1
HARNESS_BLUE = RGBColor(0, 173, 230)     # #00ADE6
HARNESS_INDIGO = RGBColor(94, 72, 232)   # #5E48E8
TD_GREEN = RGBColor(0, 166, 84)          # #00A654
DARK_GRAY = RGBColor(26, 26, 26)
LIGHT_GRAY = RGBColor(74, 85, 104)
WHITE = RGBColor(255, 255, 255)

def parse_markdown(file_path):
    """Parse markdown file and extract slides"""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Split by slide markers
    slides = content.split('---')

    parsed_slides = []
    for slide in slides:
        slide = slide.strip()
        if not slide or slide.startswith('#') and 'Harness Account Plan' in slide:
            continue

        # Extract slide number and title
        lines = slide.split('\n')
        slide_data = {
            'title': '',
            'subtitle': '',
            'content': []
        }

        current_section = None
        current_list = []
        in_table = False
        table_data = []

        for line in lines:
            line_stripped = line.strip()

            # Skip empty lines at start
            if not line_stripped and not slide_data['title']:
                continue

            # Extract slide title (## SLIDE X: Title)
            if line_stripped.startswith('## SLIDE'):
                match = re.match(r'## SLIDE \d+: (.+)', line_stripped)
                if match:
                    slide_data['title'] = match.group(1)

            # Extract subtitle (### Subtitle)
            elif line_stripped.startswith('### '):
                slide_data['subtitle'] = line_stripped[4:]
                current_section = 'subtitle'

            # Detect table
            elif '|' in line_stripped and not in_table:
                in_table = True
                table_data = [line_stripped]
            elif in_table:
                if '|' in line_stripped:
                    table_data.append(line_stripped)
                else:
                    in_table = False
                    slide_data['content'].append({
                        'type': 'table',
                        'data': table_data
                    })
                    table_data = []

            # Bullet points
            elif line_stripped.startswith('- ') or line_stripped.startswith('• '):
                bullet = line_stripped[2:].strip()
                # Handle markdown links
                bullet = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', bullet)
                slide_data['content'].append({
                    'type': 'bullet',
                    'text': bullet,
                    'level': 1
                })

            # Numbered lists
            elif re.match(r'^\d+\.\s', line_stripped):
                bullet = re.sub(r'^\d+\.\s+', '', line_stripped)
                slide_data['content'].append({
                    'type': 'bullet',
                    'text': bullet,
                    'level': 1
                })

            # Bold section headers within content
            elif line_stripped.startswith('**') and line_stripped.endswith('**'):
                header = line_stripped.strip('*')
                slide_data['content'].append({
                    'type': 'header',
                    'text': header
                })

            # Regular paragraphs
            elif line_stripped and not line_stripped.startswith('#'):
                # Skip checkmarks and emojis for cleaner slides
                if not line_stripped.startswith(('✅', '❌', '⚠️', '🔴', '🌍', '🚀')):
                    slide_data['content'].append({
                        'type': 'paragraph',
                        'text': line_stripped
                    })

        if slide_data['title']:
            parsed_slides.append(slide_data)

    return parsed_slides

def add_title_slide(prs, title, subtitle):
    """Add title slide"""
    slide_layout = prs.slide_layouts[0]  # Title slide layout
    slide = prs.slides.add_slide(slide_layout)

    title_shape = slide.shapes.title
    subtitle_shape = slide.placeholders[1]

    title_shape.text = title
    subtitle_shape.text = subtitle

    # Format title
    title_frame = title_shape.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(44)
    title_para.font.bold = True
    title_para.font.color.rgb = HARNESS_PURPLE

    # Format subtitle
    subtitle_frame = subtitle_shape.text_frame
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.font.size = Pt(24)
    subtitle_para.font.color.rgb = DARK_GRAY

def add_content_slide(prs, slide_data):
    """Add content slide"""
    # Use blank layout for more control
    slide_layout = prs.slide_layouts[5]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)

    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = slide_data['title']
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_para.font.color.rgb = DARK_GRAY

    # Add subtitle if exists
    top = Inches(1.2)
    if slide_data['subtitle']:
        subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.9), Inches(9), Inches(0.4))
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.text = slide_data['subtitle']
        subtitle_para = subtitle_frame.paragraphs[0]
        subtitle_para.font.size = Pt(20)
        subtitle_para.font.italic = True
        subtitle_para.font.color.rgb = HARNESS_INDIGO
        top = Inches(1.4)

    # Add content
    content_box = slide.shapes.add_textbox(Inches(0.5), top, Inches(9), Inches(5.5))
    content_frame = content_box.text_frame
    content_frame.word_wrap = True

    for idx, item in enumerate(slide_data['content']):
        if idx > 0:
            p = content_frame.add_paragraph()
        else:
            p = content_frame.paragraphs[0]

        if item['type'] == 'header':
            p.text = item['text']
            p.font.size = Pt(18)
            p.font.bold = True
            p.font.color.rgb = HARNESS_PURPLE
            p.space_after = Pt(6)

        elif item['type'] == 'bullet':
            p.text = item['text']
            p.level = item.get('level', 1) - 1
            p.font.size = Pt(14)
            p.font.color.rgb = DARK_GRAY
            p.space_after = Pt(4)

        elif item['type'] == 'paragraph':
            p.text = item['text']
            p.font.size = Pt(14)
            p.font.color.rgb = LIGHT_GRAY
            p.space_after = Pt(8)

        elif item['type'] == 'table':
            # Skip table in text box, will add separately
            pass

def add_table_slide(prs, slide_data):
    """Add slide with table"""
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)

    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = slide_data['title']
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_para.font.color.rgb = DARK_GRAY

    # Add subtitle
    if slide_data['subtitle']:
        subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.9), Inches(9), Inches(0.4))
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.text = slide_data['subtitle']
        subtitle_para = subtitle_frame.paragraphs[0]
        subtitle_para.font.size = Pt(20)
        subtitle_para.font.italic = True
        subtitle_para.font.color.rgb = HARNESS_INDIGO

    # Find table data
    table_data = None
    for item in slide_data['content']:
        if item['type'] == 'table':
            table_data = item['data']
            break

    if table_data:
        # Parse table
        rows = []
        for line in table_data:
            if '---' not in line:  # Skip separator line
                cells = [cell.strip() for cell in line.split('|')]
                cells = [c for c in cells if c]  # Remove empty cells
                if cells:
                    rows.append(cells)

        if rows:
            # Create table
            num_rows = len(rows)
            num_cols = len(rows[0])

            left = Inches(0.5)
            top = Inches(2.0)
            width = Inches(9)
            height = Inches(0.4) * num_rows

            table = slide.shapes.add_table(num_rows, num_cols, left, top, width, height).table

            # Fill table
            for i, row in enumerate(rows):
                for j, cell_text in enumerate(row):
                    cell = table.cell(i, j)
                    cell.text = cell_text

                    # Format header row
                    if i == 0:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = HARNESS_PURPLE
                        for paragraph in cell.text_frame.paragraphs:
                            paragraph.font.color.rgb = WHITE
                            paragraph.font.bold = True
                            paragraph.font.size = Pt(12)
                    else:
                        for paragraph in cell.text_frame.paragraphs:
                            paragraph.font.size = Pt(11)

    # Add remaining content
    content_items = [item for item in slide_data['content'] if item['type'] != 'table']
    if content_items:
        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.5), Inches(9), Inches(2.5))
        content_frame = content_box.text_frame
        content_frame.word_wrap = True

        for idx, item in enumerate(content_items):
            if idx > 0:
                p = content_frame.add_paragraph()
            else:
                p = content_frame.paragraphs[0]

            if item['type'] == 'bullet':
                p.text = item['text']
                p.level = 0
                p.font.size = Pt(13)
                p.font.color.rgb = DARK_GRAY

def main():
    """Generate PowerPoint presentation"""
    print("Parsing markdown file...")
    slides = parse_markdown('/home/user/bktest/TD-Bank-Account-Plan.md')

    print(f"Found {len(slides)} slides")

    print("Creating PowerPoint presentation...")
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # Add title slide
    add_title_slide(prs,
                   "Harness + TD Bank",
                   "Transforming Software Delivery for North America's Most Convenient Bank")

    # Add content slides
    for slide_data in slides:
        # Check if slide has table
        has_table = any(item['type'] == 'table' for item in slide_data['content'])

        if has_table:
            add_table_slide(prs, slide_data)
        else:
            add_content_slide(prs, slide_data)

        print(f"Added slide: {slide_data['title']}")

    # Save presentation
    output_path = '/home/user/bktest/TD-Bank-Account-Plan.pptx'
    prs.save(output_path)
    print(f"\nPresentation saved to: {output_path}")
    print(f"Total slides: {len(prs.slides)}")

if __name__ == '__main__':
    main()
