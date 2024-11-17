# models.py
from django.db import models

class PDFData(models.Model):
    date = models.DateField(null=True, blank=True)
    designation = models.CharField(max_length=100, blank=True)
    institution_name = models.CharField(max_length=100, blank=True)
    address = models.CharField(max_length=255, blank=True)
    subject = models.CharField(max_length=255, blank=True)
    project_name = models.CharField(max_length=100, blank=True)
    amount = models.FloatField(max_length=50, blank=True)
    proprietor_name = models.CharField(max_length=100, blank=True)
    letterhead = models.ImageField(upload_to='letterheads/', blank=True, null=True)
    signature = models.ImageField(upload_to='signatures/', blank=True, null=True)
    excel = models.FileField(upload_to='signatures/', blank=True, null=True)