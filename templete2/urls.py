from django.urls import path
from . import views

urlpatterns = [
    path('generatepdf/', views.generate_pdf_view, name='generate_pdf'),
    path('priceschedule/', views.generate_priceschedule, name='priceschedule'),
    path('quotation/', views.generate_quotation, name='quotation'),
]
