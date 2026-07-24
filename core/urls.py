from django.contrib.auth import views as auth_views
from django.urls import path

from . import views

app_name = "core"

urlpatterns = [
    path("login/", views.BrandedLoginView.as_view(), name="login"),
    path("logout/", auth_views.LogoutView.as_view(next_page="core:login"), name="logout"),
    path("", views.upload_view, name="upload"),
    path("bank-statement/", views.bank_statement_upload_view, name="bank_statement_upload"),
    path("result/<int:log_id>/", views.result_view, name="result"),
    path("toggle-passed/<int:log_id>/", views.toggle_passed_view, name="toggle_passed"),
    path("download/<int:log_id>/<str:kind>/", views.download_file_view, name="download_file"),
    path("audit-log/", views.audit_log_view, name="audit_log"),
    path("audit-log/export/", views.export_audit_log_view, name="export_audit_log"),
    path("dashboard/", views.dashboard_view, name="dashboard"),
    path("dashboard/export/failed/", views.export_failed_onoffus_view, name="export_failed_onoffus"),
    path("dashboard/export/summary/", views.export_dashboard_summary_view, name="export_dashboard_summary"),
    path("dashboard/export/days/", views.export_day_breakdown_view, name="export_day_breakdown"),
]
