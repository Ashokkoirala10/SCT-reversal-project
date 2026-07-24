from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("core", "0001_initial"),
    ]

    operations = [
        migrations.AddField(
            model_name="processinglog",
            name="prabhu_rerouted",
            field=models.IntegerField(default=0),
        ),
        migrations.AddField(
            model_name="processinglog",
            name="duplicate_skipped",
            field=models.IntegerField(default=0),
        ),
        migrations.AddField(
            model_name="processinglog",
            name="failed_reason_breakdown",
            field=models.JSONField(blank=True, default=dict),
        ),
    ]
