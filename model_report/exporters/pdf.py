# -*- coding: utf-8 -*-
import cStringIO as StringIO
from cgi import escape
from xhtml2pdf import pisa

from django.template.loader import get_template
from django.template import Context
from django.http import HttpResponse

from .base import Exporter


class PdfExporter(Exporter):

    @classmethod
    def render(cls, report, column_labels, report_rows, report_inlines):
        # where is_export is being used?
        setattr(report, 'is_export', True)
        context_dict = {
            'report': report,
            'column_labels': column_labels,
            'report_rows': report_rows,
            'report_inlines': report_inlines,
            'pagesize': 'legal landscape'
        }
        template = get_template('model_report/export_pdf.html')
        html = template.render(context_dict)
        result = StringIO.StringIO()
        pdf_encoding = 'UTF-8'

        pdf = pisa.CreatePDF(html, result, encoding=pdf_encoding)

        if not pdf.err:
            response = HttpResponse(result.getvalue(), content_type='application/pdf')
            response['Content-Disposition'] = 'attachment; filename=%s.pdf' % report.slug
        else:
            response = HttpResponse('We had some errors<pre>%s</pre>' % escape(html))

        result.close()
        return response
