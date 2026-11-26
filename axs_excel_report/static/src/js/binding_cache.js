/** @odoo-module **/
// Odoo keeps loaded views, including their Print and Action menus, in a browser
// cache. Changing an Excel template must show up on the next screen, so the view
// cache is cleared whenever a template is saved, archived or deleted.
import { rpcBus } from "@web/core/network/rpc";

const UPDATES = ["create", "write", "unlink", "web_save", "copy", "action_archive", "action_unarchive"];

rpcBus.addEventListener("RPC:RESPONSE", (ev) => {
    const params = ev.detail?.data?.params || {};
    if (params.model === "axs.excel.template" && UPDATES.includes(params.method)) {
        rpcBus.trigger("CLEAR-CACHES", "get_views");
    }
});
