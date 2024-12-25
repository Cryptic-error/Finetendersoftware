# models.py
from django.db import models


class PDFData(models.Model):
    date = models.DateField(null=True, blank=True)
    designation = models.CharField(max_length=100, blank=True)
    ourdesignation = models.CharField(max_length=100, blank=True)
    institution_name = models.CharField(max_length=100, blank=True)
    address = models.CharField(max_length=255, blank=True)
    id_no = models.CharField(max_length=255, blank=True)
    subject = models.CharField(max_length=255, blank=True)
    project_name = models.CharField(max_length=100, blank=True)
    nameofcontract = models.CharField(max_length=100, blank=True)
    amount = models.FloatField(max_length=50, blank=True)
    days = models.DecimalField(max_digits=50,decimal_places=0,blank=True)
    bidvaliditydays = models.DecimalField(default=90,max_digits=50,decimal_places=0,blank=True)
    proprietor_name = models.CharField(max_length=100, blank=True)
    attornity_name = models.CharField(max_length=100, blank=True)
    attornity_designation = models.CharField(max_length=100, blank=True)
    datedthis = models.CharField(max_length=100, blank=True,help_text="3 day of December 2024")
    letterhead = models.ImageField(upload_to='letterheads/', blank=True, null=True)
    signature = models.ImageField(upload_to='signatures/', blank=True, null=True)
    attornitysign = models.ImageField(upload_to='attornitysign/', blank=True, null=True)
    footer = models.ImageField(upload_to='footer/', blank=True, null=True)
    
    


class generatequotation(models.Model):
    institution_name = models.CharField(max_length=200, blank=True)
    designation = models.CharField(max_length=200, blank=True)
    subject = models.CharField(max_length=200, blank=True)
    # institutionname = models.CharField(max_length=200, blank=True)
    address = models.CharField(max_length=200, blank=True)
    paragraph =  models.CharField( max_length=1000,blank=True, default="We are pleased to present the pricing for the specified items in your upcoming tender for a lab setups . Here is  the estimate for your Products required ")
    notes = models.CharField(max_length=200, blank=True, default="Unit price shall include all custom duties and taxes, transportation cost tothe final destination and insurance cost.")
    amount= models.FloatField(max_length=50, blank=True)
    letterhead = models.ImageField(upload_to='letterheads/', blank=True, null=True)
    signature = models.ImageField(upload_to='signatures/', blank=True, null=True)
    excel = models.FileField(upload_to='signatures/', blank=True, null=True)
    

class priceschedule(models.Model):
    institutionname = models.CharField(max_length=200, blank=True)
    notes = models.CharField(max_length=200, blank=True, default="Unit price shall include all custom duties and taxes, transportation cost tothe final destination and insurance cost.")
    amount= models.FloatField(max_length=50, blank=True)
    letterhead = models.ImageField(upload_to='letterheads/', blank=True, null=True)
    signature = models.ImageField(upload_to='signatures/', blank=True, null=True)
    excel = models.FileField(upload_to='signatures/', blank=True, null=True)