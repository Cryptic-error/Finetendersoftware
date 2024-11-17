# forms.py
from django import forms
from .models import PDFData


class PDFForm(forms.ModelForm):
    include_selfdeclarationletter = forms.BooleanField(required=False, label="self declaration letter")
    include_warrantyletter = forms.BooleanField(required=False, label="Include warranty letter")
    include_manufactureletter = forms.BooleanField(required=False, label="manufatureletter")
    include_deliverycommitmentletter = forms.BooleanField(required=False, label="delivery commitment letter")
    include_installationletter = forms.BooleanField(required=False, label="installation letter")
    include_attorneyletter = forms.BooleanField(required=False, label="attorney letter")
    include_bidsubmissionform = forms.BooleanField(required=False, label="bidsubmission letter")
    include_technicalbid = forms.BooleanField(required=False, label="Include Technical bid")
    include_table_data = forms.BooleanField(required=False, label="Include Table Data")
    # Add other checkbox fields as needed

    class Meta:
        model = PDFData
        fields = ['date', 'designation', 'institution_name', 'address', 'subject', 'project_name', 'amount', 'proprietor_name', 'letterhead','signature','excel']
