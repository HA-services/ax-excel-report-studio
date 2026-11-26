from . import models
from . import wizard


def _uninstall_hook(*args):
    """Remove the menu actions of the reports (``(cr, registry)`` on Odoo 16, ``env`` after)."""
    from odoo import SUPERUSER_ID, api
    env = args[0] if len(args) == 1 else api.Environment(args[0], SUPERUSER_ID, {})
    env['axs.excel.template'].with_context(active_test=False).search([]).mapped('action_id').unlink()
