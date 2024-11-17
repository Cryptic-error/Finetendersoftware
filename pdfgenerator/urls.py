from django.urls import path
from . import views

urlpatterns = [
    path('generatepdf/', views.generate_pdf_view, name='generatepdf'),
]
