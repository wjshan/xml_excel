# -*-coding:utf-8 -*-
# @author: DanDan
# @contact: 454273687@qq.com
# @file: serializer.py
# @time: 2021/8/29 15:20
# @desc:
import datetime

import dateutil.parser


class SerializerAble(object):
    @property
    def tag_name(self):
        raise NotImplementedError

    @classmethod
    def from_excel(cls, instance, *args, **kwargs):
        raise NotImplementedError

    def to_excel(self, parent, *args, **kwargs):
        raise NotImplementedError

    @classmethod
    def from_xml(cls, node, *args, **kwargs):
        raise NotImplementedError

    def to_xml(self, *args, **kwargs):
        raise NotImplementedError

    @staticmethod
    def convert_python_value(value, datatype=None):
        if value == 'none' and datatype != 'str':
            return None
        if value is None:
            return None
        if not datatype:
            return value
        if datatype == 'date':
            return dateutil.parser.parse(value).date()
        if datatype == 'datetime':
            return dateutil.parser.parse(value).date()
        if datatype == 'int':
            return int(value)
        if datatype == 'float':
            return float(value)
        return value

    @staticmethod
    def convert_xml_value(value):
        if value is None:
            return 'none'
        if isinstance(value, (datetime.datetime, datetime.date)):
            return value.isoformat()
        if isinstance(value, (int, float)):
            return str(value)
        return value
