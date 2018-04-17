#!/usr/bin/python
# -*- coding: utf-8 -*-

# standard library imports
import sys
import xml.etree.ElementTree
import imp

# third party imports
import xlrd

__version__ = '1.0'
__author__ = 'wangxueyuan@navinfo.com (wangxueyuan)'
__copyright__ = "Copyright (C) 2016, NavInfo Co., Ltd."
__email__ = "wangxueyuan@navinfo.com"
__status__ = "Development"




class QueryType(object):
    UNKNOWN = 0x00
    INSERT = 0x01
    UPDATE = 0x02
    DELETE = 0x03
    SELECT = 0x04


class QueryBuilder(object):
    def __init__(self, table_name=None, type=None, match_columns=None, match_filter_columns=None, values=None,
                 filter_values=None):
        self.table_name = table_name
        self.type = type
        self.match_columns = match_columns
        self.match_filter_columns = match_filter_columns
        self.values = values
        self.filter_values = filter_values

    def build_select(self):
        value=""
        query = 'select ({filter_values});'.format(
            filter_values=self._normalize(value,self.filter_values,self.match_filter_columns)
        )
        return query

    def build(self):
        if self.type == QueryType.SELECT:
            return self.build_select()
        if self.type == QueryType.UPDATE:
            return self.build_update()
        if self.type == QueryType.DELETE:
            return self.build_delete()
        return None

    def _normalize(self, s, filter_values,match_filter_columns):
        if s is None:
            return ''
        for i in range(len(match_filter_columns)):
            if match_filter_columns[i]=="生产需求":
                s+="".join(filter_values[i])
                # s="wang"
        return s


class XLS2SQL(object):
    def __init__(self, filename=None, sheet=None, table_name=None, query_type=None, columns=None, match_columns=None,
                 filter_columns=None, match_filter_columns=None, header_index=0, verbose=False):
        self.book = xlrd.open_workbook(filename)
        self.sheet = self.book.sheet_by_name(sheet)
        self.table_name = table_name
        self.query_type = query_type
        self.columns = columns
        self.match_columns = match_columns
        self.filter_columns = filter_columns
        self.match_filter_columns = match_filter_columns
        self.header_index = header_index
        self.values_indexes = []
        self.filter_indexes = []
        self.verbose = verbose

        self._initialize_filters()

    def run(self):
        for i in range(self.sheet.nrows):
            if i == self.header_index:
                continue
            for j in self.values_indexes:
                if self.sheet.row_values(i)[j]=="值域检查":
                    filter_values = [self.sheet.row_values(i)[j] for j in self.filter_indexes]
                    builder = QueryBuilder(
                        table_name=self.table_name,
                        type=self.query_type,
                        match_columns=self.match_columns,
                        match_filter_columns=self.filter_columns,
                        values="",
                        filter_values=filter_values
                    )
                    print(builder.build())



    def _load_indexes(self, ref, values, res):
        for col in ref:
            found = False
            for k in range(len(values)):
                if col == values[k]:
                    found = True
                    res.append(k)
                    break
            if not found:
                raise Exception('missing column %s' % col)

    def _initialize_filters(self):
        for i in range(self.sheet.nrows):
            if i == self.header_index:
                row = self.sheet.row_values(i)
                self._load_indexes(self.columns, row, self.values_indexes)
                self._load_indexes(self.filter_columns, row, self.filter_indexes)

    @staticmethod
    def from_xml(filename):
        doc = xml.etree.ElementTree.parse(filename).getroot()

        filename = None
        sheet = None
        table_name = None
        query_type = None
        columns = []
        match_columns = []
        filter_columns = []
        match_filter_columns = []
        header_index = 0

        for elem in doc:
            if elem.tag == 'filename':
                filename = elem.text;
            elif elem.tag == 'sheet':
                sheet = elem.text
            elif elem.tag == 'table':
                table_name = elem.attrib.get('name', None)
            elif elem.tag == 'query':
                t = elem.attrib.get('type', None)
                if t.upper() == 'UPDATE':
                    query_type = QueryType.UPDATE
                elif t.upper() == 'SELECT':
                    query_type = QueryType.SELECT
                elif t.upper() == 'DELETE':
                    query_type = QueryType.DELETE
            elif elem.tag == 'column':
                columns.append(elem.attrib.get('excel', None))#columns的表头
                match_columns.append(elem.attrib.get('table', None))#对应字段名
            elif elem.tag == 'filter':
                filter_columns.append(elem.attrib.get('excel', None))#filter的表头名
                match_filter_columns.append(elem.attrib.get('table', None))
            elif elem.tag == 'header':
                header_index = int(elem.attrib.get('index', 0))

        xls2sql = XLS2SQL(
            filename=filename,
            sheet=sheet,
            table_name=table_name,
            query_type=query_type,
            columns=columns,
            match_columns=match_columns,
            filter_columns=filter_columns,
            match_filter_columns=match_filter_columns,
            header_index=header_index
        )

        return xls2sql

if __name__ == '__main__':
    xls2sql = XLS2SQL.from_xml('0014-add-rules-20180302.xml')
    xls2sql.run()
