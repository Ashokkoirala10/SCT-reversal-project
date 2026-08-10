from django.db import models
from django.utils import timezone


def upload_path(instance, filename):
    return f"uploads/{instance.id or 'tmp'}_{filename}"


def output_path(instance, filename):
    return f"outputs/{filename}"


class ProcessingLog(models.Model):
    """Central audit trail: one row per file processed."""

    STATUS_SUCCESS = "SUCCESS"
    STATUS_FAILED = "FAILED"
    STATUS_CHOICES = [
        (STATUS_SUCCESS, "Success"),
        (STATUS_FAILED, "Failed"),
    ]

    uploaded_file = models.FileField(upload_to="uploads/")
    uploaded_filename = models.CharField(max_length=255)
    generated_file = models.FileField(upload_to="outputs/", blank=True, null=True)
    generated_filename = models.CharField(max_length=255, blank=True)

    uploaded_by = models.CharField(max_length=150, blank=True, help_text="Optional: who ran this")

    status = models.CharField(max_length=10, choices=STATUS_CHOICES, default=STATUS_SUCCESS)
    error_message = models.TextField(blank=True)

    # A generated file only shows up in the central "shared reports" list
    # and only counts toward the analytics dashboard once its owner (or an
    # admin) has reviewed it and explicitly marked it as passed. Until
    # then it only shows up in that user's own "My activity" list.
    passed = models.BooleanField(default=False)
    passed_by = models.CharField(max_length=150, blank=True, default="")
    passed_at = models.DateTimeField(null=True, blank=True)

    total_rows = models.IntegerField(default=0)
    success_count = models.IntegerField(default=0)
    success_amount = models.FloatField(default=0.0)
    failed_total = models.IntegerField(default=0)
    failed_insufficient_funds = models.IntegerField(default=0)
    failed_kept = models.IntegerField(default=0)
    # Amount/charge summed across every FAILED row (kept + insufficient
    # funds), for reconciliation. Charge already reflects the "NCHL
    # transactions carry the charge inside the amount, so charge is
    # treated as 0" business rule.
    failed_amount = models.FloatField(default=0.0)
    failed_charge = models.FloatField(default=0.0)
    # Amount/charge summed only across the rows kept in the "failed"
    # sheet (failed_kept), needed to cleanly subtract out "failed but
    # credited" rows once a bank statement has been checked.
    failed_kept_amount = models.FloatField(default=0.0)
    failed_kept_charge = models.FloatField(default=0.0)
    reversal_total = models.IntegerField(default=0)
    reversal_manual_kept = models.IntegerField(default=0)
    reversal_manual_amount = models.FloatField(default=0.0)
    reversal_manual_charge = models.FloatField(default=0.0)
    # REVERSAL rows that were NOT a manual reversal (i.e. already reversed
    # automatically by the switch) — not written to the output file, but
    # tracked here for daily reconciliation.
    reversal_system_count = models.IntegerField(default=0)
    reversal_system_amount = models.FloatField(default=0.0)
    reversal_system_charge = models.FloatField(default=0.0)
    # Subset of reversal_system_* that is On-Us (Global-to-Global or
    # Prabhu-to-Prabhu) or NCHL-routed — written to their own sheet
    # (see ONUS_SYSTEM_REVERSAL_SHEET_NAME in core/services.py) instead of
    # only being counted.
    reversal_system_onus_count = models.IntegerField(default=0)
    reversal_system_onus_amount = models.FloatField(default=0.0)
    reversal_system_onus_charge = models.FloatField(default=0.0)
    coop_count = models.IntegerField(default=0)
    coop_member_count = models.IntegerField(default=0)
    imeremit_count = models.IntegerField(default=0)
    cityremit_count = models.IntegerField(default=0)
    timeout_count = models.IntegerField(default=0)
    timeout_amount = models.FloatField(default=0.0)

    # Prabhu Bank rows that got the special debit account instead of the
    # normal clearing account.
    prabhu_rerouted = models.IntegerField(default=0)
    # Manual-reversal rows skipped because they already appeared in the
    # previous generated reversal file (double-reversal prevention).
    duplicate_skipped = models.IntegerField(default=0)
    # Rows skipped because their Network Reference Id already showed up in
    # an *earlier upload's source file* — catches the "downloaded an
    # overlapping time window" case (e.g. yesterday's file already covered
    # today's midnight-to-download-time slice).
    duplicate_source_skipped = models.IntegerField(default=0)
    # How long process_ibft_file() (dedup + reconciliation + workbook
    # write) took for this upload, in milliseconds — shown on the
    # analytics dashboard's day-by-day breakdown.
    processing_duration_ms = models.IntegerField(default=0)
    # The latest "Transaction Date" found in the uploaded source file —
    # i.e. how far into the day this export's data actually reaches.
    # Shown on the dashboard's day-by-day breakdown instead of/alongside
    # wall-clock processing time.
    data_through_at = models.DateTimeField(null=True, blank=True)
    # Manual-reversal rows whose Debtor Bank was neither Global IME Bank nor
    # Prabhu Bank — still processed (using the standard account), but
    # flagged here so an admin can review and, if needed, add proper
    # handling for that bank.
    unrecognized_debtor_bank_rows = models.IntegerField(default=0)
    unrecognized_debtor_banks = models.JSONField(default=list, blank=True)
    # {reason: count} breakdown of every FAILED row's Source Message, for
    # the analytics dashboard.
    failed_reason_breakdown = models.JSONField(default=dict, blank=True)

    # On-Us = Debtor Bank is Global IME Bank (our own bank); Off-Us = any
    # other bank. Tracked separately (with their own reason breakdowns) so
    # the dashboard can show where failures are actually coming from.
    failed_onus_count = models.IntegerField(default=0)
    failed_onus_amount = models.FloatField(default=0.0)
    failed_offus_count = models.IntegerField(default=0)
    failed_offus_amount = models.FloatField(default=0.0)
    failed_reason_breakdown_onus = models.JSONField(default=dict, blank=True)
    failed_reason_breakdown_offus = models.JSONField(default=dict, blank=True)

    # Manual-reversal rows whose Debtor Bank is Prabhu Bank — these are
    # written to their own dedicated "prabhu" sheet in the generated file
    # instead of being lumped into coop/imeremit/cityremit, so Prabhu
    # reversals can be worked (and bank-statement-checked) separately.
    prabhu_reversal_count = models.IntegerField(default=0)
    prabhu_reversal_amount = models.FloatField(default=0.0)

    # --- Bank statement cross-check (optional, applied after generation) ---
    BANK_TYPE_GLOBAL = "global"
    BANK_TYPE_PRABHU = "prabhu"
    BANK_TYPE_BOTH = "both"
    BANK_TYPE_CHOICES = [
        (BANK_TYPE_GLOBAL, "Global IME Bank"),
        (BANK_TYPE_PRABHU, "Prabhu Bank"),
        (BANK_TYPE_BOTH, "Global IME Bank + Prabhu Bank"),
    ]
    bank_statement_checked = models.BooleanField(default=False)
    bank_statement_filename = models.CharField(max_length=255, blank=True, default="")
    # Which bank the most recent check's statement(s) were for. Prabhu
    # statements may be up to 4 separate daily exports, dumped into one
    # combined file before processing (see combine_bank_statement_files());
    # bank_statement_filename then holds a comma-separated list of the
    # original filenames.
    bank_statement_type = models.CharField(max_length=10, choices=BANK_TYPE_CHOICES, blank=True, default="")
    bank_statement_checked_by = models.CharField(max_length=150, blank=True, default="")
    bank_statement_checked_at = models.DateTimeField(null=True, blank=True)
    # FAILED rows whose Network Reference Id showed up in the bank
    # statement (money moved despite a FAILED status) — red-flagged on
    # the "failed" sheet and copied into "Failed but Credited".
    failed_credited_count = models.IntegerField(default=0)
    failed_credited_amount = models.FloatField(default=0.0)
    failed_credited_charge = models.FloatField(default=0.0)
    # Manual-reversal rows (coop/imeremit/cityremit) whose underlying
    # credit was found already reversed out in the bank statement —
    # red-flagged in place to prevent a double reversal.
    already_reversed_count = models.IntegerField(default=0)
    already_reversed_amount = models.FloatField(default=0.0)
    already_reversed_charge = models.FloatField(default=0.0)
    # Subset of the above caught specifically by the NCHL direct-reference-id
    # (CR-alongside-DR) check — see is_already_debited_nchl() in services.py.
    already_reversed_nchl_count = models.IntegerField(default=0)
    already_reversed_nchl_amount = models.FloatField(default=0.0)
    # Subset of the above caught specifically by the Khalti settlement-
    # account (KHALTI_SETTL) check — see is_already_debited_khalti() in
    # services.py.
    already_reversed_khalti_count = models.IntegerField(default=0)
    already_reversed_khalti_amount = models.FloatField(default=0.0)
    # Subset of already_reversed_* caught by the On-Us "already successful"
    # (duplicate DR) verification on the manual-reversal sheets themselves
    # — see has_duplicate_dr() in core/services.py. Catches an On-Us
    # (Global-to-Global or Prabhu-to-Prabhu) row that landed in
    # coop/imeremit/cityremit/prabhu by mistake even though it already
    # completed successfully end-to-end.
    already_reversed_onus_success_count = models.IntegerField(default=0)
    already_reversed_onus_success_amount = models.FloatField(default=0.0)
    # Rows on the On-Us/NCHL system reversal sheet found to have already
    # succeeded BEFORE the system reversal was issued — see
    # is_onus_already_success() in core/services.py.
    onus_system_reversal_flagged_count = models.IntegerField(default=0)
    onus_system_reversal_flagged_amount = models.FloatField(default=0.0)
    # On-Us (Debtor Bank & Creditor Bank both our own Global IME Bank) FAILED
    # rows found to have a duplicate DR in the statement — both legs of the
    # transfer actually completed, so these are genuinely successful and do
    # NOT need a manual reversal (see has_duplicate_dr() in services.py).
    # Green-flagged on "failed" and copied into "Already Success (OnUs)".
    onus_already_success_count = models.IntegerField(default=0)
    onus_already_success_amount = models.FloatField(default=0.0)
    onus_already_success_charge = models.FloatField(default=0.0)

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-created_at"]

    def __str__(self):
        # created_at is stored in UTC internally (USE_TZ=True); always show
        # it converted to the configured local time zone (Asia/Kathmandu) —
        # otherwise dropdowns/labels built from this (e.g. the bank
        # statement "reversal file" picker) show a time that's hours off
        # from what the file was actually generated at.
        local_dt = timezone.localtime(self.created_at) if self.created_at else None
        when = local_dt.strftime("%Y-%m-%d %H:%M") if local_dt else "—"
        return f"{self.uploaded_filename} -> {self.generated_filename or '(failed)'} @ {when}"


class SeenNetworkReferenceId(models.Model):
    """Every Network Reference Id that has ever been processed out of an
    uploaded ibft-transaction file, regardless of its Overall Status.

    Consecutive daily exports typically overlap (e.g. "midnight to
    whenever I downloaded" every day, so the hours between midnight and
    yesterday's download time get pulled twice). Before a new upload is
    processed, its rows are checked against this table and any row whose
    Network Reference Id is already here is skipped — see
    `duplicate_source_skipped` on ProcessingLog / process_ibft_file()'s
    `seen_reference_ids` parameter.
    """

    ref_id = models.CharField(max_length=64, unique=True, db_index=True)
    source_log = models.ForeignKey(
        ProcessingLog, on_delete=models.CASCADE, related_name="seen_refs"
    )
    first_seen_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.ref_id


class MemberAggregatorStat(models.Model):
    """Per (Member Name, Aggregator) breakdown of Success / Failed / Manual
    Reversal counts + amounts, for one processed upload.

    Written once per upload (see process_ibft_file()'s
    `member_aggregator_breakdown` on ProcessingStats), and rolled up across
    every passed file for the Member & Aggregator report on the dashboard.
    """

    log = models.ForeignKey(ProcessingLog, on_delete=models.CASCADE, related_name="member_aggregator_stats")
    member_name = models.CharField(max_length=255, blank=True, default="")
    aggregator = models.CharField(max_length=255, blank=True, default="")

    success_count = models.IntegerField(default=0)
    success_amount = models.FloatField(default=0.0)
    failed_count = models.IntegerField(default=0)
    failed_amount = models.FloatField(default=0.0)
    reversal_count = models.IntegerField(default=0)
    reversal_amount = models.FloatField(default=0.0)

    class Meta:
        indexes = [models.Index(fields=["member_name", "aggregator"])]

    def __str__(self):
        return f"{self.member_name} / {self.aggregator} (log {self.log_id})"


class BankAccount(models.Model):
    """Configurable Debtor-Bank -> Debit-Account mapping, editable from the
    Django admin (Admin > Core > Bank accounts).

    Global IME Bank and Prabhu Bank ship as the two default rows (seeded by
    migration 0014), matching what used to be hardcoded constants in
    core/services.py: editing a row's Debit Account Number here changes
    which account is placed in a reversal row's "Debit Account Number"
    column for any transaction whose Debtor Bank matches this row's
    Keyword — no code change or deploy needed. Add a brand-new row the same
    way to support a bank that isn't Global IME or Prabhu at all; tick
    "Is own bank" if it should also count toward On-Us determination
    (both Debtor Bank and Creditor Bank being one of our own banks).
    """

    bank_name = models.CharField(
        max_length=100,
        help_text="Display name, e.g. 'Global IME Bank'.",
    )
    keyword = models.CharField(
        max_length=50,
        unique=True,
        help_text=(
            "Upper-cased, case-insensitive substring matched against the "
            "'Debtor Bank' (and 'Creditor Bank') column, e.g. 'GLOBAL' "
            "matches 'Global IME Bank'. Keep this short and specific."
        ),
    )
    debit_account_number = models.CharField(
        max_length=50,
        help_text="Placed in 'Debit Account Number' for any reversal row whose Debtor Bank matches this Keyword.",
    )
    is_own_bank = models.BooleanField(
        default=False,
        help_text=(
            "Check for Global IME Bank / Prabhu Bank and any other bank we "
            "operate — this bank then counts toward On-Us determination "
            "(Debtor Bank and Creditor Bank both being one of our own "
            "banks), which drives the On-Us \"already successful\" checks."
        ),
    )
    is_active = models.BooleanField(
        default=True,
        help_text="Untick to stop matching this bank without deleting the row.",
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["bank_name"]

    def __str__(self):
        return f"{self.bank_name} ({self.keyword})"

    def save(self, *args, **kwargs):
        self.keyword = (self.keyword or "").strip().upper()
        super().save(*args, **kwargs)


def _clear_bank_account_cache(**kwargs):
    # Lazy import: services.py has no top-level dependency on models.py, and
    # this keeps it that way — only reached once Django has fully loaded.
    from . import services

    services.bank_accounts_cache_clear()


models.signals.post_save.connect(_clear_bank_account_cache, sender=BankAccount)
models.signals.post_delete.connect(_clear_bank_account_cache, sender=BankAccount)
