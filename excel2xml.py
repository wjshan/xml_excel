from openpyxl import load_workbook
from xml_excel.workbook import WorkbookSerializer
from lxml import etree


def excel2xml(excel_file, export_file, read_range=None):
    wb = load_workbook(excel_file)
    serializer = WorkbookSerializer.from_excel(wb,read_range=read_range)
    workbook_tag, style_tag = serializer.to_xml()
    with open(export_file, 'wb') as f:
        f.write(etree.tostring(workbook_tag, encoding='utf-8', pretty_print=True))
    with open(f"{export_file}.style", 'wb') as f:
        f.write(etree.tostring(style_tag, encoding='utf-8', pretty_print=True))


if __name__ == '__main__':
    import sys

    args = sys.argv[1:]
    excel_file_name = args[0]
    if len(args) > 1:
        export_file_name = args[1]
    else:
        export_file_name = f"{excel_file_name}.xml"
    excel2xml(excel_file_name, export_file_name)
    print(export_file_name)

    # excel2xml('/Users/dandan/Downloads/劳务分包造价统计.xlsx', '/Users/dandan/Downloads/劳务分包造价统计.xlsx.xml')
