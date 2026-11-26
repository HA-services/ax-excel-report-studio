from odoo import api, models


class IrActions(models.Model):
    _inherit = 'ir.actions.actions'

    @api.model
    def get_bindings(self, model_name):
        """Hide the Excel report entries that are not available in the current company."""
        result = super().get_bindings(model_name)
        ids = [a['id'] for actions in result.values() for a in actions]
        if not ids:
            return result
        templates = self.env['axs.excel.template'].sudo().search([('action_id', 'in', ids)])
        hidden = {t.action_id.id for t in templates if not t._is_available()}
        if not hidden:
            return result
        filtered = {}
        for binding_type, actions in result.items():
            kept = [a for a in actions if a['id'] not in hidden]
            if kept:
                filtered[binding_type] = kept
        return filtered
