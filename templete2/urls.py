from django.urls import path
from . import views

urlpatterns = [
    path('generatepdf/', views.generate_pdf_view, name='temp_generate_pdf'),
    path('priceschedule/', views.generate_priceschedule, name='temp_priceschedule'),
    path('quotation/', views.generate_quotation, name='temp_quotation'),
    path('quotation-word/', views.generate_quotation_docs, name='temp_quotation_docs'),
    path('generate-word/', views.generate_word, name='temp_generate_word'),
    
    # other paths
]
