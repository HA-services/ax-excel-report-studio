"""API differences between Odoo 16 and Odoo 20."""
import base64


def can_read(Model):
    if hasattr(Model, 'has_access'):  # Odoo 18+
        return Model.has_access('read')
    return Model.check_access_rights('read', raise_exception=False)


def check_read(records):
    """Raise AccessError unless the current user can read ``records`` (rights and rules)."""
    if hasattr(records, 'check_access'):  # Odoo 18+
        records.check_access('read')
    else:
        records.check_access_rights('read')
        records.check_access_rule('read')


def user_has_groups(env, groups):
    """True when the current user belongs to ``groups`` (comma separated xml ids, ``!`` for exclusion)."""
    if env.su:
        return True
    if hasattr(env.user, 'has_groups'):  # Odoo 18+
        return env.user.has_groups(groups)
    return env['base'].user_has_groups(groups)


def binary_read(value):
    """Raw bytes of a Binary field value (base64 up to Odoo 19, BinaryValue on Odoo 20)."""
    if not value:
        return b''
    if hasattr(value, 'content'):
        return value.content
    return base64.b64decode(value)


def binary_write(raw):
    """Value to store raw bytes in a Binary field."""
    try:
        from odoo.tools import BinaryBytes
    except ImportError:
        return base64.b64encode(raw)
    return BinaryBytes(raw)
