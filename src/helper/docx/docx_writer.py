#!/usr/bin/env python3

import json
import time
import pprint

from docx.shared import Pt, Cm, Inches, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_BREAK
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION, WD_ORIENT

from helper.logger import *
from helper.docx.docx_util import *

VALIGN = {'TOP': WD_CELL_VERTICAL_ALIGNMENT.TOP, 'MIDDLE': WD_CELL_VERTICAL_ALIGNMENT.CENTER, 'BOTTOM': WD_CELL_VERTICAL_ALIGNMENT.BOTTOM}
HALIGN = {'LEFT': WD_ALIGN_PARAGRAPH.LEFT, 'CENTER': WD_ALIGN_PARAGRAPH.CENTER, 'RIGHT': WD_ALIGN_PARAGRAPH.RIGHT, 'JUSTIFY': WD_ALIGN_PARAGRAPH.JUSTIFY}

def merged_cell_width(row, col, start_row, start_col, merges, column_widths):
    cell_width = 0
    for merge in merges:
        if merge['startRowIndex'] == (row + start_row) and merge['startColumnIndex'] == (col + start_col):
            for c in range(col, merge['endColumnIndex'] - start_col):
                cell_width = cell_width + column_widths[c]

    if cell_width == 0:
        return column_widths[col]
    else:
        return cell_width

def set_header(doc, section, header_first, header_odd, header_even, actual_width, linked_to_previous=False):
    first_page_header = section.first_page_header
    odd_page_header = section.header
    even_page_header = section.even_page_header

    section.first_page_header.is_linked_to_previous = linked_to_previous
    section.header.is_linked_to_previous = linked_to_previous
    section.even_page_header.is_linked_to_previous = linked_to_previous

    if len(first_page_header.tables) == 0:
        if header_first is not None: insert_content(header_first, doc, actual_width, container=first_page_header, cell=None)

    if len(odd_page_header.tables) == 0:
        if header_odd is not None: insert_content(header_odd, doc, actual_width, container=odd_page_header, cell=None)

    if len(even_page_header.tables) == 0:
        if header_even is not None: insert_content(header_even, doc, actual_width, container=even_page_header, cell=None)

def set_footer(doc, section, footer_first, footer_odd, footer_even, actual_width, linked_to_previous=False):
    first_page_footer = section.first_page_footer
    odd_page_footer = section.footer
    even_page_footer = section.even_page_footer

    section.first_page_footer.is_linked_to_previous = linked_to_previous
    section.footer.is_linked_to_previous = linked_to_previous
    section.even_page_footer.is_linked_to_previous = linked_to_previous

    if len(first_page_footer.tables) == 0:
        if footer_first is not None: insert_content(footer_first, doc, actual_width, container=first_page_footer, cell=None)

    if len(odd_page_footer.tables) == 0:
        if footer_odd is not None: insert_content(footer_odd, doc, actual_width, container=odd_page_footer, cell=None)

    if len(even_page_footer.tables) == 0:
        if footer_even is not None: insert_content(footer_even, doc, actual_width, container=even_page_footer, cell=None)

def add_section(doc, section_data, section_spec, use_existing=False):
    if section_spec['break'] == 'CONTINUOUS':
        # if it is the only section, do not add, just get the last (current) section
        if use_existing:
            section = doc.sections[-1]
        else:
            section = doc.add_section(WD_SECTION.CONTINUOUS)
    else:
        # if it is the only section, do not add, just get the last (current) section
        if use_existing:
            section = doc.sections[-1]
        else:
            section = doc.add_section(WD_SECTION.NEW_PAGE)

    if section_spec['orient'] == 'LANDSCAPE':
        section.orient = WD_ORIENT.LANDSCAPE
    else:
        section.orient = WD_ORIENT.PORTRAIT

    section.page_width = Inches(section_spec['page_width'])
    section.page_height = Inches(section_spec['page_height'])
    section.left_margin = Inches(section_spec['left_margin'])
    section.right_margin = Inches(section_spec['right_margin'])
    section.top_margin = Inches(section_spec['top_margin'])
    section.bottom_margin = Inches(section_spec['bottom_margin'])
    section.header_distance = Inches(section_spec['header_distance'])
    section.footer_distance = Inches(section_spec['footer_distance'])
    section.gutter = Inches(section_spec['gutter'])
    section.different_first_page_header_footer = section_data['different-first-page-header-footer']

    # get the actual width
    actual_width = section.page_width.inches - section.left_margin.inches - section.right_margin.inches - section.gutter.inches

    # set header if it is not set already
    set_header(doc, section, section_data['header-first'], section_data['header-odd'], section_data['header-even'], actual_width)

    # set footer if it is not set already
    set_footer(doc, section, section_data['footer-first'], section_data['footer-odd'], section_data['footer-even'], actual_width)

    return section

def render_cell(doc, cell, cell_data, width, r, c, start_row, start_col, merge_data, column_widths):
    cell.width = Inches(width)
    paragraph = cell.paragraphs[0]

    # if there is a note, see if it is a JSON, it may contain style, page-numering, new-page, keep-with-next directive etc.
    note_json = {}
    if 'note' in cell_data:
        try:
            note_json = json.loads(cell_data['note'])
        except json.JSONDecodeError:
            pass

    # process new-page
    if 'new-page' in note_json:
        # return the cell location so that the page break can be rendered later
        pf = paragraph.paragraph_format
        pf.page_break_before = True

    # process keep-with-next
    if 'keep-with-next' in note_json:
        # return the cell location so that the page break can be rendered later
        pf = paragraph.paragraph_format
        pf.keep_with_next = True

    # do some special processing if the cell_data is {}
    if cell_data == {} or 'effectiveFormat' not in cell_data:
        return

    text_format = cell_data['effectiveFormat']['textFormat']
    effective_format = cell_data['effectiveFormat']

    # alignments
    cell.vertical_alignment = VALIGN[effective_format['verticalAlignment']]
    if 'horizontalAlignment' in effective_format:
        paragraph.alignment = HALIGN[effective_format['horizontalAlignment']]

    # background color
    bgcolor = cell_data['effectiveFormat']['backgroundColor']
    if bgcolor != {}:
        red = int(bgcolor['red'] * 255) if 'red' in bgcolor else 0
        green = int(bgcolor['green'] * 255) if 'green' in bgcolor else 0
        blue = int(bgcolor['blue'] * 255) if 'blue' in bgcolor else 0
        set_cell_bgcolor(cell, RGBColor(red, green, blue))

    # text-rotation
    if 'textRotation' in effective_format:
        text_rotation = effective_format['textRotation']
        rotate_text(cell, 'btLr')

    # borders
    if 'borders' in cell_data['effectiveFormat']:
        borders = cell_data['effectiveFormat']['borders']
        set_cell_border(cell, top=ooxml_border_from_gsheet_border(borders, 'top'), bottom=ooxml_border_from_gsheet_border(borders, 'bottom'), start=ooxml_border_from_gsheet_border(borders, 'left'), end=ooxml_border_from_gsheet_border(borders, 'right'))

    # cell can be merged, so we need width after merge (in In)
    cell_width = merged_cell_width(r, c, start_row, start_col, merge_data, column_widths)

    # images
    if 'userEnteredValue' in cell_data:
        userEnteredValue = cell_data['userEnteredValue']
        if 'image' in userEnteredValue:
            image = userEnteredValue['image']
            run = paragraph.add_run()

            # even now the width may exceed actual cell width, we need to adjust for that
            # determine cell_width based on merge scenario
            dpi_x = 150 if image['dpi'][0] == 0 else image['dpi'][0]
            dpi_y = 150 if image['dpi'][1] == 0 else image['dpi'][1]
            image_width = image['width'] / dpi_x
            image_height = image['height'] / dpi_y
            if image_width > cell_width:
                adjust_ratio = (cell_width / image_width)
                # keep a padding of 0.1 inch
                image_width = cell_width - 0.2
                image_height = image_height * adjust_ratio

            run.add_picture(image['path'], height=Inches(image_height), width=Inches(image_width))

    # before rendering cell, see if it embeds another worksheet
    if 'contents' in cell_data:
        table = insert_content(cell_data['contents'], doc, cell_width, container=None, cell=cell)
        polish_table(table)
        return

    # texts
    if 'formattedValue' not in cell_data:
        return

    text = cell_data['formattedValue']

    # process notes
    # note specifies style
    if 'style' in note_json:
        paragraph.add_run(text)
        paragraph.style = note_json['style']
        return

    # note specifies page numbering
    if 'page-number' in note_json:
        append_page_number_with_pages(paragraph)
        #append_page_number_only(paragraph)
        paragraph.style = note_json['page-number']
        return

    # finally cell content, add runs
    if 'textFormatRuns' in cell_data:
        text_runs = cell_data['textFormatRuns']
        # split the text into run-texts
        run_texts = []
        for i in range(len(text_runs) - 1, -1, -1):
            text_run = text_runs[i]
            if 'startIndex' in text_run:
                run_texts.insert(0, text[text_run['startIndex']:])
                text = text[:text_run['startIndex']]
            else:
                run_texts.insert(0, text)

        # now render runs
        for i in range(0, len(text_runs)):
            # get formatting
            format = text_runs[i]['format']

            run = paragraph.add_run(run_texts[i])
            set_character_style(run, {**text_format, **format})
    else:
        run = paragraph.add_run(text)
        set_character_style(run, text_format)

def insert_content(data, doc, container_width, container=None, cell=None):
    start_time = int(round(time.time() * 1000))
    current_time = int(round(time.time() * 1000))
    if not container: debug('.. inserting contents')
    last_time = current_time

    start_row, start_col = data['sheets'][0]['data'][0]['startRow'], data['sheets'][0]['data'][0]['startColumn']

    # calculate table dimension
    table_rows = data['sheets'][0]['properties']['gridProperties']['rowCount'] - start_row
    table_cols = data['sheets'][0]['properties']['gridProperties']['columnCount'] - start_col

    merge_data = {}
    if 'merges' in data['sheets'][0]:
        merge_data = data['sheets'][0]['merges']

    # create the table
    if container is not None:
        table = container.add_table(table_rows, table_cols, Pt(container_width))

	# table to be added inside a cell
    elif cell is not None:
		# insert the table in the very first paragraph of the cell
        # cell._element.clear_content()
        cell.paragraphs[0].style = 'Calibri-2-Gray8'
        table = cell.add_table(table_rows, table_cols)
    else:
        table = doc.add_table(table_rows, table_cols)

    # resize columns as per data
    column_data = data['sheets'][0]['data'][0]['columnMetadata']
    total_width = sum(x['pixelSize'] for x in column_data)
    column_widths = [ (x['pixelSize'] * container_width / total_width) for x in column_data ]

    # if the table had too many columns, use a style where there is smaller left, right margin
    if len(column_widths) > 10:
        table.style = 'PlainTable'

    last_time = current_time

    # populate cells
    total_rows = len(data['sheets'][0]['data'][0]['rowData'])
    i = 0
    current_time = int(round(time.time() * 1000))
    if not container: info('  .. rendering cell for {0} rows'.format(total_rows))
    last_time = current_time

    row_data = data['sheets'][0]['data'][0]['rowData']
    for r in range(0, len(row_data)):
        if 'values' in row_data[r]:
            row = table.row_cells(r)
            row_values = row_data[r]['values']

            for c in range(0, len(row_values)):
                render_cell(doc, row[c], row_values[c], column_widths[c], r, c, start_row, start_col, merge_data, column_widths)

            if r % 100 == 0:
                current_time = int(round(time.time() * 1000))
                if not container: info('  .... cell rendered for {0}/{1} rows : {2} ms'.format(r, total_rows, current_time - last_time))
                last_time = current_time

    current_time = int(round(time.time() * 1000))
    if not container: info('  .. rendering cell complete for {0} rows : {1} ms\n'.format(total_rows, current_time - start_time))
    last_time = current_time

    # merge cells according to data
    if not container: info('  .. merging cells'.format(current_time - last_time))
    for m in merge_data:
        startRowIndex = m['startRowIndex'] - start_row
        endRowIndex = m['endRowIndex'] - start_row - 1
        startColumnIndex = m['startColumnIndex'] - start_col
        endColumnIndex = m['endColumnIndex'] - start_col - 1
        starting_cell = table.cell(startRowIndex, startColumnIndex)
        # all cells within the merge range need to have the same border as the first cell
        for r in range(startRowIndex, endRowIndex + 1):
            for c in range(startColumnIndex, endColumnIndex + 1):
                if (r, c) != (startRowIndex, startColumnIndex):
                    to_cell = table.cell(r, c)
                    copy_cell_border(starting_cell, to_cell)

        ending_cell = table.cell(endRowIndex, endColumnIndex)
        starting_cell.merge(ending_cell)

    current_time = int(round(time.time() * 1000))
    if not container: info('  .. cells merged : {0} ms\n'.format(current_time - last_time))
    last_time = current_time

    if not container: info('.. content insertion completed : {0} ms\n'.format(current_time - start_time))

    return table
