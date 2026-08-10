from django.contrib import admin

from .models import BankAccount, ProcessingLog


@admin.register(BankAccount)
class BankAccountAdmin(admin.ModelAdmin):
    list_display = (
        "id", "bank_name", "keyword", "debit_account_number",
        "is_own_bank", "is_active", "updated_at",
    )
    list_filter = ("is_own_bank", "is_active")
    search_fields = ("bank_name", "keyword", "debit_account_number")


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
        "prabhu_reversal_count",
        "unrecognized_debtor_bank_rows",
        "duplicate_skipped",
        "created_at",
    )
    list_filter = ("status", "passed", "created_at")
    search_fields = ("uploaded_filename", "generated_filename", "uploaded_by")
    readonly_fields = [f.name for f in ProcessingLog._meta.fields]

# MemberAggregatorStat is intentionally NOT registered here — it's a
# per-upload, per-(member, aggregator) row (hundreds/thousands of rows per
# file) that isn't meaningful to browse/manage one-by-one in Django admin.
# It's already fully exposed, rolled up and filterable, on the dashboard's
# Member-wise / Aggregator-wise report tabs (see core/views.py) and via
# their Excel exports, so a Django admin listing for it is unnecessary.
