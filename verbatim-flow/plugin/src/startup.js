window.Asc.plugin.init = function () {
    // Setting default values for the first row of cells
    Api.GetCell(0, 0).Value = "1AC";
    Api.GetCell(1, 0).Value = "1NC";
    Api.GetCell(2, 0).Value = "2AC";
    Api.GetCell(3, 0).Value = "Block";
    Api.GetCell(4, 0).Value = "1AR";
    Api.GetCell(5, 0).Value = "2NR";
    Api.GetCell(6, 0).Value = "2AR";
}