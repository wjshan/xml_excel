# -*-coding:utf-8 -*-
# @author: DanDan
# @contact: 454273687@qq.com
# @file: workbook.py
# @time: 2021/8/29 15:50
# @desc:
from .serializer import SerializerAble
from openpyxl import Workbook
from lxml import etree
from .sheet import SheetSerializer
from openpyxl.workbook.defined_name import DefinedName


class WorkbookSerializer(SerializerAble):
    tag_name = 'WorkBook'

    def __init__(self):
        self.sheets = []

    @classmethod
    def from_excel(cls, work_book: Workbook, **kwargs):
        instance = cls()
        for work_sheet in work_book:
            instance.sheets.append(SheetSerializer.from_excel(work_sheet, **kwargs))
        return instance

    def to_excel(self, *args, **kwargs):
        wb = Workbook()
        wb.remove(wb.active)
        for index, sheet in enumerate(self.sheets):
            sheet.to_excel(wb, index=index)
        return wb

    @classmethod
    def from_xml(cls, node, *args, **kwargs):
        self = cls()
        style_node = kwargs.get('style_node')
        if not len(style_node):
            style_node = node
        for sheet_node in node.xpath(f'//{SheetSerializer.tag_name}'):
            self.sheets.append(SheetSerializer.from_xml(sheet_node, style_node=style_node))
        return self

    def to_xml(self, *args, **kwargs):
        workbook_node = etree.Element(self.tag_name)
        style_tag = etree.Element('styles')
        for work_sheet in self.sheets:
            sheet_tag, style_nodes = work_sheet.to_xml()
            workbook_node.append(sheet_tag)
            for style in style_nodes:
                style_tag.append(style)
        return workbook_node, style_tag
