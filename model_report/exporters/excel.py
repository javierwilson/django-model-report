# -*- coding: utf-8 -*-
from xlwt import Workbook, easyxf

from django.http import HttpResponse

from model_report import arial10
from .base import Exporter


class FitSheetWrapper(object):
    """Try to fit columns to max size of any entry.
    To use, wrap this around a worksheet returned from the
    workbook's add_sheet method, like follows:

        sheet = FitSheetWrapper(book.add_sheet(sheet_name))

    The worksheet interface remains the same: this is a drop-in wrapper
    for auto-sizing columns.
    """
    def __init__(self, sheet):
        self.sheet = sheet
        self.widths = dict()
        self.heights = dict()

    def write(self, r, c, label='', *args, **kwargs):
        self.sheet.write(r, c, label, *args, **kwargs)
        self.sheet.row(r).collapse = True
        bold = False
        if args:
            style = args[0]
            bold = str(style.font.bold) in ('1', 'true', 'True')
        width = int(arial10.fitwidth(label, bold))
        if width > self.widths.get(c, 0):
            self.widths[c] = width
            self.sheet.col(c).width = width

        height = int(arial10.fitheight(label, bold))
        if height > self.heights.get(r, 0):
            self.heights[r] = height
            self.sheet.row(r).height = height

    def __getattr__(self, attr):
        return getattr(self.sheet, attr)


class ExcelExporter(Exporter):

    def write_rows(self, column_labels, report_rows, report_inlines=None):
        for index, x in enumerate(column_labels):
            self.sheet1.write(self.row_index, index, u'%s' % x, self.stylebold)
        self.row_index += 1
        for g, rows in report_rows:
            if g:
                self.sheet1.write(self.row_index, 0, u'%s' % g, self.stylebold)
                self.row_index += 1
            for row in list(rows):
                if row.is_value():
                    for index, x in enumerate(row):
                        if isinstance(x.value, (list, tuple)):
                            xvalue = ''.join(['%s\n' % v for v in x.value])
                        else:
                            xvalue = x.text()
                        self.sheet1.write(self.row_index, index, xvalue, self.stylevalue)
                    self.row_index += 1

                    if report_inlines:
                        for inline in report_inlines:

                            inline_context = inline.get_render_context({}, by_row=row)
                            self.write_rows(inline_context['column_labels'], inline_context['report_rows'])

                elif row.is_caption:
                    for index, x in enumerate(row):
                        if not isinstance(x, (unicode, str)):
                            self.sheet1.write(self.row_index, index, x.text(), self.stylebold)
                        else:
                            self.sheet1.write(self.row_index, index, x, self.stylebold)
                    self.row_index += 1
                elif row.is_total:
                    for index, x in enumerate(row):
                        self.sheet1.write(self.row_index, index, x.text(), self.stylebold)
                        self.sheet1.write(self.row_index + 1, index, ' ')
                    self.row_index += 2


    def render(self, report, column_labels, report_rows, report_inlines):
        self.row_index = 0
        self.sheet1 = FitSheetWrapper(self.book.add_sheet(report.get_title()[:20]))
        self.write_rows(column_labels, report_rows, report_inlines)

        response = HttpResponse(content_type="application/ms-excel")
        response['Content-Disposition'] = 'attachment; filename=%s.xls' % report.slug
        self.book.save(response)
        return response

    def __init__(self):
        self.stylebold = easyxf('font: bold true; alignment:')
        self.stylevalue = easyxf('alignment: horizontal left, vertical top;')
        self.book = Workbook(encoding='utf-8')
