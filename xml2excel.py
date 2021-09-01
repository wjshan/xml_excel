from utils.xml_excel.workbook import WorkbookSerializer
from lxml import etree


def _xml2excel(data_tag, style_tag):
    serializer = WorkbookSerializer.from_xml(data_tag, style_node=style_tag)
    wb = serializer.to_excel()
    return wb


def xml2excel(data_xml, style_xml):
    work_book_tag = etree.fromstring(data_xml)
    style_tag = etree.fromstring(style_xml)
    return _xml2excel(work_book_tag, style_tag)
