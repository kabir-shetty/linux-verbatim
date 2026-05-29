function getToolbarItems() {
    console.log()
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