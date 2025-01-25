# forms.py
from django import forms
from .models import PDFData,generatequotation


class PDFForm(forms.ModelForm):
    include_selfdeclarationletter = forms.BooleanField(required=False, label="self declaration letter")
    include_warrantyletter = forms.BooleanField(required=False, label="Include warranty letter")
    include_manufactureletter = forms.BooleanField(required=False, label="manufatureletter")
    include_deliverycommitmentletter = forms.BooleanField(required=False, label="delivery commitment letter")
    include_installationletter = forms.BooleanField(required=False, label="installation letter")
    include_attorneyletter = forms.BooleanField(required=False, label="attorney letter")
    include_bidsubmissionform = forms.BooleanField(required=False, label="bidsubmission letter")
    include_technicalbid = forms.BooleanField(required=False, label="Include Technical bid")
    include_quotationandprice = forms.BooleanField(required=False, label="include quotation and price")
    include_table_data = forms.BooleanField(required=False, label="Include Table Data")
    include_pricebid = forms.BooleanField(required=False, label="Include Price bid")
    # Add other checkbox fields as needed

    class Meta:
        model = PDFData
        fields = ['date', 'designation','ourdesignation','id_no', 'institution_name', 'days','bidvaliditydays','address', 'subject','attornity_name','attornity_designation','attornitysign', 'project_name', 'amount', 'proprietor_name', 'letterhead','signature']

class quotationform(forms.ModelForm):
    
    include_table_data = forms.BooleanField(required=False, label="Include Table Data")
    
    class Meta:
        model = generatequotation
        fields = "__all__"


class pricescheduleform(forms.ModelForm):
    
    include_table_data = forms.BooleanField(required=False, label="Include Table Data")
    
    class Meta:
        model = generatequotation
        fields = "__all__"
    
# class quotationform(forms.ModelForm):
#     class Meta:
#         model = generatequotation
#         fields = ['institution_name', 'excel']