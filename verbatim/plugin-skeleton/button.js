window.Asc.plugin.executeMethod("InputText", ["button.js loaded"]);

function getToolbarItems() {
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

window.Asc.plugin.init = function () {
    this.executeMethod("AddToolbarMenuItem", [getToolbarItems()]);
}