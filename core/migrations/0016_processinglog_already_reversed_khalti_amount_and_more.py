# Generated manually, following the same pattern as
# 0010_processinglog_already_reversed_nchl_amount_and_more.py

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0015_seed_bank_accounts'),
    ]

    operations = [
        migrations.AddField(
            model_name='processinglog',
            name='already_reversed_khalti_amount',
            field=models.FloatField(default=0.0),
        ),
        migrations.AddField(
            model_name='processinglog',
            name='already_reversed_khalti_count',
            field=models.IntegerField(default=0),
        ),
    ]
