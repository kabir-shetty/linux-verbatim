window.Asc.plugin.init = function () {
    function getToolbarItems() {
        window.Asc.plugin.executeMethod("InputText", ["geToolbarItems ran"]);
        let items = {
          guid: window.Asc.plugin.info.guid,
          tabs: [{
            id: "tab_1",
            text: "Insert options",
            items: []
          }]
        };
    
        return items;
    };

    window.Asc.plugin.executeMethod("InputText", ["button.js loaded"]);
    this.executeMethod("AddToolbarMenuItem", [getToolbarItems()]);
}