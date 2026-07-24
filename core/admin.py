from django.contrib import admin

from .models import ProcessingLog


@admin.register(ProcessingLog)
class ProcessingLogAdmin(admin.ModelAdmin):
    list_display = (
        "id",
        "uploaded_filename",
        "generated_filename",
        "status",
        "passed",
        "uploaded_by",
        "reversal_manual_kept",
        "failed_kept",
        "timeout_count",
        "prabhu_rerouted",
        "unrecognized_debtor_bank_rows",
        "duplicate_skipped",
        "created_at",
    )
    list_filter = ("status", "passed", "created_at")
    search_fields = ("uploaded_filename", "generated_filename", "uploaded_by")
    readonly_fields = [f.name for f in ProcessingLog._meta.fields]
