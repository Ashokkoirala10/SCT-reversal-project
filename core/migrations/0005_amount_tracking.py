from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("core", "0004_passed_workflow"),
    ]

    operations = [
        migrations.AddField(
            model_name="processinglog",
            name="success_amount",
            field=models.FloatField(default=0.0),
        ),
        migrations.AddField(
            model_name="processinglog",
            name="failed_amount",
            field=models.FloatField(default=0.0),
        ),
        migrations.AddField(
            model_name="processinglog",
            name="failed_charge",
            field=models.FloatField(default=0.0),
        ),
        migrations.AddField(
            model_name="processinglog",
            name="reversal_manual_amount",
            field=models.FloatField(default=0.0),
        ),
        migrations.AddField(
            model_name="processinglog",
            name="reversal_manual_charge",
            field=models.FloatField(default=0.0),
        ),
        migrations.AddField(
            model_name="processinglog",
            name="reversal_system_count",
            field=models.IntegerField(default=0),
        ),
        migrations.AddField(
            model_name="processinglog",
            name="reversal_system_amount",
            field=models.FloatField(default=0.0),
        ),
        migrations.AddField(
            model_name="processinglog",
            name="reversal_system_charge",
            field=models.FloatField(default=0.0),
        ),
        migrations.AddField(
            model_name="processinglog",
            name="timeout_amount",
            field=models.FloatField(default=0.0),
        ),
    ]
