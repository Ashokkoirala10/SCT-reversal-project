from django.contrib import admin

from .models import BankAccount, MemberAggregatorStat, ProcessingLog


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


@admin.register(MemberAggregatorStat)
class MemberAggregatorStatAdmin(admin.ModelAdmin):
    list_display = (
        "id", "log", "member_name", "aggregator",
        "success_count", "failed_count", "reversal_count",
    )
    list_filter = ("aggregator",)
    search_fields = ("member_name", "aggregator")
    readonly_fields = [f.name for f in MemberAggregatorStat._meta.fields]
