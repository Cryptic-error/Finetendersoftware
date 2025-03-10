# Generated by Django 5.1.2 on 2025-01-14 07:04

import django.db.models.deletion
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('templete2', '0003_generatequotation_terms1_generatequotation_terms2_and_more'),
    ]

    operations = [
        migrations.AddField(
            model_name='generatequotation',
            name='merged_pdf',
            field=models.FileField(blank=True, null=True, upload_to='merged_pdfs/'),
        ),
        migrations.CreateModel(
            name='UploadedFile',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('file', models.FileField(upload_to='uploaded_pdfs/')),
                ('row_name', models.CharField(blank=True, max_length=200, null=True)),
                ('file_type', models.CharField(choices=[('Financial Docs', 'Financial Docs'), ('BOQ', 'BOQ'), ('Catalogue', 'Catalogue'), ('CE', 'CE'), ('ISO', 'ISO'), ('VAT', 'VAT')], max_length=20)),
                ('uploaded_at', models.DateTimeField(auto_now_add=True)),
                ('quotation', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='templete2.generatequotation')),
            ],
        ),
    ]
