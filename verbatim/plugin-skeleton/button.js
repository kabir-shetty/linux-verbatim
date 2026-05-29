window.Asc.plugin.init = function () {
    function getToolbarItems() {
        window.Asc.plugin.executeMethod("InputText", ["geToolbarItems ran"]);
        let items = {
          guid: window.Asc.plugin.info.guid,
          tabs: [{
            id: "tab_1",
            text: "Insert options",
            items: [
                {
                    id: "button_1",
                    type: "button",
                    text: "this shit pmo"
                }
            ]
          }]
        };
    
        return items;
    };

    window.Asc.plugin.executeMethod("InputText", ["button.js loaded"]);
    window.Asc.plugin.executeMethod("AddToolbarMenuItem", [getToolbarItems()]);
}