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


class BankStatementUploadForm(forms.Form):
    target_log = forms.ModelChoiceField(
        queryset=ProcessingLog.objects.none(),
        label="Reversal file to check this statement against",
        empty_label="Select a generated reversal file\u2026",
    )
    statement_file = forms.FileField(
        label="Bank statement (.csv or .xlsx)",
        widget=forms.ClearableFileInput(attrs={"accept": ".csv,.xlsx"}),
    )

    def __init__(self, *args, user=None, **kwargs):
        super().__init__(*args, **kwargs)
        qs = ProcessingLog.objects.filter(
            status=ProcessingLog.STATUS_SUCCESS, generated_file__isnull=False
        ).order_by("-created_at")
        if user is not None and not user.is_staff:
            qs = qs.filter(uploaded_by=user.username)
        self.fields["target_log"].queryset = qs

    def clean_statement_file(self):
        f = self.cleaned_data["statement_file"]
        name = f.name.lower()
        if not (name.endswith(".csv") or name.endswith(".xlsx")):
            raise forms.ValidationError(
                "Please upload a .csv or .xlsx bank statement export."
            )
        return f
