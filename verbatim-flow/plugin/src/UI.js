function getToolbarItems() {
  let items = {
    guid: window.Asc.plugin.info.guid,
    tabs: [{
      id: "ribbon",
      text: "Verbatim",
      items: [
          {
              id: "button_1",
              type: "button",
              text: "idk lowk"
          }
      ]
    }]
  };

  return items;
};

window.Asc.plugin.init = function () {
    window.Asc.plugin.executeMethod("AddToolbarMenuItem", [getToolbarItems()]);
}