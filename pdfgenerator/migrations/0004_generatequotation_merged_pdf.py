# Generated by Django 5.1.2 on 2025-01-02 11:04

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('pdfgenerator', '0003_uploadedfile_file_type_alter_uploadedfile_row_name'),
    ]

    operations = [
        migrations.AddField(
            model_name='generatequotation',
            name='merged_pdf',
            field=models.FileField(blank=True, null=True, upload_to='merged_pdfs/'),
        ),
    ]
