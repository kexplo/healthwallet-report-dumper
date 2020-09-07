import os
from typing import List, Set
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

import xlsxwriter


def dump_ui_xml() -> str:
    os.system('adb shell uiautomator dump')
    os.system('adb pull /sdcard/window_dump.xml')
    os.system('adb shell rm /sdcard/window_dump.xml')
    return os.path.abspath('./window_dump.xml')


def send_scroll_down() -> None:
    os.system('adb shell input swipe 500 1000 500 500')


def get_header(ui_xml_path: str) -> List[str]:
    tree = ET.parse(ui_xml_path)
    root = tree.getroot()
    header = root.find('.//node[@class="android.widget.ListView"]/../node[1]')
    assert isinstance(header, Element)
    header_texts = [node.get('text') for node in header.findall('node')]
    # for debug
    print(' | '.join(header_texts))
    return header_texts


def get_listview_items(ui_xml_path: str) -> List[List[str]]:
    tree = ET.parse(ui_xml_path)
    root = tree.getroot()
    listview = root.find('.//node[@class="android.widget.ListView"]')
    assert isinstance(listview, Element)
    ## UI structure ----------------------------
    ############################################
    # ListView
    # |- LinearLayout        (a Row)
    #    |- LinearLayout
    #       |- View
    #       |- LinearLayout  (Content of a Row)
    #          |- TextView   (a Column)
    #          |- View       (a Column separator?)
    #          |- TextView   (a Column)
    #          |- ...
    # |- LinearLayout        (a Row)
    # |- ...
    #-------------------------------------------

    # There is two type of LinearLayout (a Row)
    ## Type 1
    #
    # |- LinearLayout        (a Row)
    #    |- LinearLayout
    #       |- View
    #       |- LinearLayout  (Content of a Row)
    #          |- TextView   (a Column)
    #          |- View       (a Column separator?)
    #          |- TextView   (a Column)
    #          |- ...
    #
    ## Type 2
    #
    # |- LinearLayout        (a Row)
    #    |- LinearLayout
    #       |- View       (a Column separator?)
    #       |- TextView   (a Column)
    #       |- TextView   (a Column)
    #       |- View       (a Column separator?)
    columns = []
    for item in listview:
        row = item.find('./node[1]')
        texts = [
            column.get('text')
            for column in row.findall('.//node[@class="android.widget.TextView"]')  # noqa: E501
        ]  # type: List[str]
        columns.append(texts)
        # for debug
        print(' | '.join(texts))
        print('-' * 80)
    return columns


def dump() -> List[List[str]]:
    # get_column_headers('window_dump.xml')
    keys = set()  # type: Set[str]
    report = []   # type: List[List[str]]
    header = []  # type: List[str]

    while True:
        if not header:
            header = get_header(dump_ui_xml())
            report.append(header)
        columns = get_listview_items(dump_ui_xml())  # type: List[List[str]]
        key_columns = []
        for column in columns:
            key_column = column[0]
            key_columns.append(key_column)
            if key_column not in keys:
                report.append(column)
        if keys.issuperset(key_columns):
            break
        keys.update(key_columns)
        send_scroll_down()
    return report


def write_report_to_xlsx(report: List[List[str]], filename='report.xlsx'):
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    row = 0
    col = 0
    for row_data in report:
        worksheet.write_row(row, col, row_data)
        row += 1
    workbook.close()


if __name__ == '__main__':
    write_report_to_xlsx(dump())
