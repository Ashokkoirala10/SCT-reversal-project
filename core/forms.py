from django import forms

from .models import ProcessingLog


class UploadForm(forms.Form):
    ibft_file = forms.FileField(
        label="IBFT transaction file",
        widget=forms.ClearableFileInput(attrs={"accept": ".xlsx"}),
    )

    def clean_ibft_file(self):
        f = self.cleaned_data["ibft_file"]
        if not f.name.lower().endswith(".xlsx"):
            raise forms.ValidationError("Please upload an .xlsx file.")
        return f


class MultipleFileInput(forms.ClearableFileInput):
    allow_multiple_selected = True

    def value_from_datadict(self, data, files, name):
        # Django 5.0+ handles this automatically when allow_multiple_selected
        # is set, but this project targets Django 4.2+, where
        # ClearableFileInput.value_from_datadict() still only ever returns a
        # single file (files.get(name)) no matter how many were selected in
        # the browser. Explicitly pulling every file for this field name via
        # getlist() is what actually lets more than one file through (e.g.
        # up to 4 Prabhu Bank statement exports selected at once).
        if hasattr(files, "getlist"):
            uploads = files.getlist(name)
            if uploads:
                return uploads
        return files.get(name)


class MultipleFileField(forms.FileField):
    """A FileField whose widget accepts multiple files at once, returning a
    list of UploadedFile objects in cleaned_data instead of just one."""

    def __init__(self, *args, **kwargs):
        kwargs.setdefault("widget", MultipleFileInput(attrs={"multiple": True}))
        super().__init__(*args, **kwargs)

    def clean(self, data, initial=None):
        single_file_clean = super().clean
        if isinstance(data, (list, tuple)):
            return [single_file_clean(d, initial) for d in data]
        return [single_file_clean(data, initial)] if data else []


class BankStatementUploadForm(forms.Form):
    target_log = forms.ModelChoiceField(
        queryset=ProcessingLog.objects.none(),
        label="Reversal file to check this statement against",
        empty_label="Select a generated reversal file\u2026",
    )
    global_statement_files = MultipleFileField(
        label="Global IME Bank statement (.csv or .xlsx)",
        required=False,
        widget=MultipleFileInput(attrs={"accept": ".csv,.xlsx", "multiple": True}),
        help_text="Usually a single file.",
    )
    prabhu_statement_files = MultipleFileField(
        label="Prabhu Bank statement(s) (.csv or .xlsx)",
        required=False,
        widget=MultipleFileInput(attrs={"accept": ".csv,.xlsx", "multiple": True}),
        help_text="Up to 4 files (e.g. separate daily exports).",
    )

    def __init__(self, *args, user=None, **kwargs):
        super().__init__(*args, **kwargs)
        qs = ProcessingLog.objects.filter(
            status=ProcessingLog.STATUS_SUCCESS, generated_file__isnull=False
        ).order_by("-created_at")
        if user is not None and not user.is_staff:
            qs = qs.filter(uploaded_by=user.username)
        self.fields["target_log"].queryset = qs

    def clean(self):
        cleaned = super().clean()
        global_files = cleaned.get("global_statement_files") or []
        prabhu_files = cleaned.get("prabhu_statement_files") or []

        if not global_files and not prabhu_files:
            self.add_error(
                "global_statement_files",
                "Please upload at least one bank statement file (Global and/or Prabhu).",
            )
            return cleaned

        if len(global_files) > 1:
            self.add_error(
                "global_statement_files",
                "Global IME Bank statement upload supports only one file at a time.",
            )
            return cleaned

        if len(prabhu_files) > 4:
            self.add_error("prabhu_statement_files", "Please upload at most 4 Prabhu Bank statement files.")
            return cleaned

        for f in global_files + prabhu_files:
            name = f.name.lower()
            if not (name.endswith(".csv") or name.endswith(".xlsx")):
                self.add_error(
                    "global_statement_files" if f in global_files else "prabhu_statement_files",
                    f"'{f.name}' is not a .csv or .xlsx file.",
                )
                return cleaned

        cleaned["global_statement_files"] = global_files
        cleaned["prabhu_statement_files"] = prabhu_files
        return cleaned
