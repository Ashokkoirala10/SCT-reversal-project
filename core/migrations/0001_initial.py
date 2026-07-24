from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = []

    operations = [
        migrations.CreateModel(
            name="ProcessingLog",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("uploaded_file", models.FileField(upload_to="uploads/")),
                ("uploaded_filename", models.CharField(max_length=255)),
                ("generated_file", models.FileField(blank=True, null=True, upload_to="outputs/")),
                ("generated_filename", models.CharField(blank=True, max_length=255)),
                ("uploaded_by", models.CharField(blank=True, help_text="Optional: who ran this", max_length=150)),
                (
                    "status",
                    models.CharField(
                        choices=[("SUCCESS", "Success"), ("FAILED", "Failed")],
                        default="SUCCESS",
                        max_length=10,
                    ),
                ),
                ("error_message", models.TextField(blank=True)),
                ("total_rows", models.IntegerField(default=0)),
                ("success_count", models.IntegerField(default=0)),
                ("failed_total", models.IntegerField(default=0)),
                ("failed_insufficient_funds", models.IntegerField(default=0)),
                ("failed_kept", models.IntegerField(default=0)),
                ("reversal_total", models.IntegerField(default=0)),
                ("reversal_manual_kept", models.IntegerField(default=0)),
                ("coop_count", models.IntegerField(default=0)),
                ("coop_member_count", models.IntegerField(default=0)),
                ("imeremit_count", models.IntegerField(default=0)),
                ("cityremit_count", models.IntegerField(default=0)),
                ("timeout_count", models.IntegerField(default=0)),
                ("created_at", models.DateTimeField(auto_now_add=True)),
            ],
            options={
                "ordering": ["-created_at"],
            },
        ),
    ]
