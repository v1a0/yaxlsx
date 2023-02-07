from base import XlsxCell


def safe_str(value):
    if isinstance(value, XlsxCell):
        return value
    if isinstance(value, float):
        if value % 1 == 0:
            return str(int(value))
    if isinstance(value, str):
        return f'"{value}"'
    return str(value)


def format_safe_round(value):
    if isinstance(value, float):
        if value % 1 == 0:
            return int(value)

    else:
        try:
            return int(value)
        except (ValueError, TypeError):
            pass

    return value

