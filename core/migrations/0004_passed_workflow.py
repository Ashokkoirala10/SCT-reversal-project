from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("core", "0003_seed_users"),
    ]

    operations = [
        migrations.AddField(
            model_name="processinglog",
            name="passed",
            field=models.BooleanField(default=False),
        ),
        migrations.AddField(
            model_name="processinglog",
            name="passed_by",
            field=models.CharField(blank=True, default="", max_length=150),
        ),
        migrations.AddField(
            model_name="processinglog",
            name="passed_at",
            field=models.DateTimeField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name="processinglog",
            name="unrecognized_debtor_bank_rows",
            field=models.IntegerField(default=0),
        ),
        migrations.AddField(
            model_name="processinglog",
            name="unrecognized_debtor_banks",
            field=models.JSONField(blank=True, default=list),
        ),
    ]
