from django import template

register = template.Library()


@register.filter
def npr(value):
    """Format a Rupee amount compactly using the Nepali/Indian numbering
    system (Lakh = 1,00,000 / Crore = 1,00,00,000) so large, ever-growing
    totals don't blow out the width of a dashboard stat card.

    Full precision is always still available via the element's `title`
    tooltip (set separately in the template with |floatformat:2) — this
    filter is only for the big bold number itself.
    """
    try:
        v = float(value or 0)
    except (TypeError, ValueError):
        return value

    sign = "-" if v < 0 else ""
    v = abs(v)

    if v >= 1_00_00_000:
        return f"{sign}{v / 1_00_00_000:.2f} Cr"
    if v >= 1_00_000:
        return f"{sign}{v / 1_00_000:.2f} L"
    if v >= 1_000:
        return f"{sign}{v:,.0f}"
    return f"{sign}{v:,.2f}"
