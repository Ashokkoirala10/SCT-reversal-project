from django.conf import settings
from django.db import migrations


def create_hardcoded_users(apps, schema_editor):
    # Using the real (non-historical) User model here is deliberate: we
    # need password hashing (set_password) which the historical model
    # doesn't give us, and the auth User model's core fields
    # (username/password/is_staff/is_superuser/is_active) are stable.
    from django.contrib.auth.models import User

    # Regular operator account — can upload files and generate reversals,
    # but does not see the central Audit Log / Analytics Dashboard.
    user, _ = User.objects.get_or_create(
        username="ashok.koirala",
        defaults={"is_staff": False, "is_active": True},
    )
    user.set_password("ashok@123")
    user.is_staff = False
    user.is_active = True
    user.save()

    # Admin account — sees everything, including the central Audit Log and
    # the analytics dashboard.
    admin, _ = User.objects.get_or_create(
        username="admin",
        defaults={"is_staff": True, "is_superuser": True, "is_active": True},
    )
    admin.set_password("admin@123")
    admin.is_staff = True
    admin.is_superuser = True
    admin.is_active = True
    admin.save()


def remove_hardcoded_users(apps, schema_editor):
    from django.contrib.auth.models import User

    User.objects.filter(username__in=["ashok.koirala", "admin"]).delete()


class Migration(migrations.Migration):

    dependencies = [
        ("core", "0002_logging_fields"),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.RunPython(create_hardcoded_users, remove_hardcoded_users),
    ]
