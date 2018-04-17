#!/usr/bin/python
# -*- coding: utf-8 -*-

# standard library imports
import sys
import xml.etree.ElementTree
import imp

# third party imports
import xlrd

__version__ = '1.0'
__author__ = 'pavelsimo@navinfo.com (Pavel Simo)'
__copyright__ = "Copyright (C) 2016, NavInfo Co., Ltd."
__maintainer__ = "Pavel Simo"
__email__ = "pavelsimo@navinfo.com"
__status__ = "Development"




class QueryType(object):
    UNKNOWN = 0x00
    INSERT = 0x01
    UPDATE = 0x02
    DELETE = 0x03


class QueryBuilder(object):
    def __init__(self, table_name=None, type=None, match_columns=None, match_filter_columns=None, values=None,
                 filter_values=None):
        self.table_name = table_name
        self.type = type
        self.match_columns = match_columns
        self.match_filter_columns = match_filter_columns
        self.values = values
        self.filter_values = filter_values

    def build_insert(self):
        query = 'INSERT INTO {table_name}({fields}) VALUES ({values});'.format(
            table_name=self.table_name,
            fields=','.join(self.match_columns),
            values=','.join([self._normalize(value) for value in self.values])
        )
        return query

    def build_delete(self):
        filters = ["%s=%s" % (key, self._normalize(value)) for key, value in
                   zip(self.match_filter_columns, self.filter_values)]
        query = 'DELETE FROM {table_name} WHERE {filters};'.format(
            table_name=self.table_name,
            filters='AND '.join(filters)
        )
        return query

    def build_update(self):
        assignments = ["%s=%s" % (key, self._normalize(value)) for key, value in
                       zip(self.match_columns, self.values)]
        filters = ["%s=%s" % (key, self._normalize(value)) for key, value in
                   zip(self.match_filter_columns, self.filter_values)]

        query = 'UPDATE {table_name} SET {assignments} WHERE {filters};'.format(
            table_name=self.table_name,
            assignments=','.join(assignments),
            filters='AND '.join(filters)
        )
        return query

    def build(self):
        if self.type == QueryType.INSERT:
            return self.build_insert()
        if self.type == QueryType.UPDATE:
            return self.build_update()
        if self.type == QueryType.DELETE:
            return self.build_delete()
        return None

    def _normalize(self, s, to_upper=True, scape_quotes=True, strip_whitespaces=True, strip_linesep=True):
        if s is None:
            return ''
        is_str = False
        if isinstance(s, str):
            is_str = True
        if not is_str:
            s = str(s)
        if strip_linesep:
            s = s.replace('\n', ' ')
            s = s.replace('\r', ' ')
        if scape_quotes:
            s = s.replace("'", "''")
        if to_upper:
            s = s.upper()
        if strip_whitespaces:
            s = s.strip()
        if is_str:
            s = "'%s'" % s
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

        self._initialize_cols()

    def run(self):
        for i in range(self.sheet.nrows):
            if i == self.header_index:
                continue
            values = [self.sheet.row_values(i)[j] for j in self.values_indexes]
            filter_values = [self.sheet.row_values(i)[j] for j in self.filter_indexes]

            builder = QueryBuilder(
                table_name=self.table_name,
                type=self.query_type,
                match_columns=self.match_columns,
                match_filter_columns=self.match_filter_columns,
                values=values,
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

    def _initialize_cols(self):
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
                elif t.upper() == 'INSERT':
                    query_type = QueryType.INSERT
                elif t.upper() == 'DELETE':
                    query_type = QueryType.DELETE
            elif elem.tag == 'column':
                columns.append(elem.attrib.get('excel', None))
                match_columns.append(elem.attrib.get('table', None))
            elif elem.tag == 'filter':
                filter_columns.append(elem.attrib.get('excel', None))
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
    xls2sql = XLS2SQL.from_xml('configs/0014-add-rules-20180302.xml')
    xls2sql.run()
