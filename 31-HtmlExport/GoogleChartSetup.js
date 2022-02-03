    // Load the Visualization API and the corechart package.
    google.charts.load('current', {'packages': ['corechart', 'geochart'] });

    // Set a callback to run when the Google Visualization API is loaded.
    google.charts.setOnLoadCallback(drawCharts);

    function getDataFromTable(indexes) {
        // Create the data table.
        var data = new google.visualization.DataTable();
        var dataTypes = [];
        var n = 0;

        // read data types from the thead and define columns in the google.visualization.DataTable
        $("table#my-table thead tr th").each(function (i, elem) {
            if (indexes.includes(i)) {
                var dt = $(elem).data("datatype");
                dataTypes[n++] = dt;
                data.addColumn(dt, elem.innerHTML);
            }
        });

    // read the data from the tbody and insert it into the table
    let rows = [];
        $("table#my-table tbody tr").each(function (i, tableRow) {
            var row = [];
            var colIx = 0;
            indexes.forEach(ix => {
                let dataType = dataTypes[colIx++];
                var cell = $(tableRow).children().eq(ix);
                if (dataType == "string") {
                    row.push($(cell).html());
                }
                else if (dataType == "number") {
                    row.push(parseFloat($(cell).data("value")));
                }
                else if (dataType == "datetime") {
                    row.push(new Date(parseFloat($(cell).data("value"))));
                }
            });
            rows.push(row);
        });
        data.addRows(rows);
        return data;
    }

    // Callback that creates and populates a data table,
    // instantiates the pie chart, passes in the data and
    // draws it.
    function drawCharts() {
        var dt = getDataFromTable([0, 1, 3]);
        drawBarChart(dt, "FX rates");
    }

    function drawBarChart(data, title){
        var options = {
        'title': title,
        'width': 500,
        'height': 300,
        'is3D': true
        };
    var chart2 = new google.visualization.BarChart(document.getElementById('bar-chart'));
    chart2.draw(data, options);
