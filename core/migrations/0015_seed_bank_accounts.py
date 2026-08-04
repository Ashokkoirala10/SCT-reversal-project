from django.db import migrations


# Matches the previously hardcoded constants in core/services.py
# (GLOBAL_BANK_KEYWORD/GLOBAL_BANK_DEBIT_ACCOUNT and
# PRABHU_BANK_KEYWORD/PRABHU_BANK_DEBIT_ACCOUNT) — seeding these as real
# rows means the admin can now edit or add to them, while behavior on
# upgrade is completely unchanged until someone actually does.
DEFAULT_BANK_ACCOUNTS = [
    {
        "bank_name": "Global IME Bank",
        "keyword": "GLOBAL",
        "debit_account_number": "0002335524115",
        "is_own_bank": True,
        "is_active": True,
    },
    {
        "bank_name": "Prabhu Bank",
        "keyword": "PRABHU",
        "debit_account_number": "99901170130555",
        "is_own_bank": True,
        "is_active": True,
    },
]


def seed_bank_accounts(apps, schema_editor):
    BankAccount = apps.get_model("core", "BankAccount")
    for defaults in DEFAULT_BANK_ACCOUNTS:
        BankAccount.objects.get_or_create(keyword=defaults["keyword"], defaults=defaults)


def remove_seeded_bank_accounts(apps, schema_editor):
    BankAccount = apps.get_model("core", "BankAccount")
    BankAccount.objects.filter(
        keyword__in=[defaults["keyword"] for defaults in DEFAULT_BANK_ACCOUNTS]
    ).delete()


class Migration(migrations.Migration):

    dependencies = [
        ("core", "0014_bankaccount_and_more"),
    ]

    operations = [
        migrations.RunPython(seed_bank_accounts, remove_seeded_bank_accounts),
    ]
