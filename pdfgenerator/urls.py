from django.urls import path
from . import views

urlpatterns = [
    path('generatepdf/', views.generate_pdf_view, name='generate_pdf'),
    path('priceschedule/', views.generate_priceschedule, name='priceschedule'),
    path('quotation/', views.generate_quotation, name='quotation'),
    path('quotation2/', views.generate_quotation2, name='quotation2'),
    path('upload/', views.upload, name='upload'),
    path('upload_pdfs/<int:quotation_id>/', views.upload_pdfs, name='upload_pdf'),
    path('generate-word/', views.generate_word, name='generate-word'),
    path('quotation-doc/', views.quotation_docs, name='quotation-word'),
    path('quotation-doc2/', views.quotation_docs2, name='quotation-word2'),
]
