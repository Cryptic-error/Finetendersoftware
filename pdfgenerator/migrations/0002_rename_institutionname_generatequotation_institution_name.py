# Generated by Django 5.1.2 on 2024-12-12 08:10

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('pdfgenerator', '0001_initial'),
    ]

    operations = [
        migrations.RenameField(
            model_name='generatequotation',
            old_name='institutionname',
            new_name='institution_name',
        ),
    ]