# -*- coding: utf-8 -*-
try:
    from django.conf.urls.defaults import *
except ImportError:
    from django.conf.urls import url, include

from django.conf import settings
from django.conf.urls.static import static

from django.contrib import admin
admin.autodiscover()

from model_report import report
report.autodiscover()


urlpatterns = [
    url(r'^admin/', include(admin.site.urls)),
    url(r'', include('model_report.urls')),
] + static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
