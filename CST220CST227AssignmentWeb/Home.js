(function () {
    "use strict";

    //Citations: regression.js from Tom Alexander https://github.com/Tom-Alexander/regression-js
    //           simpson's rule and integration from Dennis Chapligin https://github.com/akashihi/simpson

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            //Load sample Velocity vs Time data and generate the chart of said data
            loadSampleData();

            // Add a click event handler for the regression button.
            $('#regression-button').click(doRegression);

            //Show/hide the order input box if polynomial is chosen
            $('input[type=radio]').on('change', function () {
                if ($('#polynomial').is(':checked')) {
                    $('#order').show();
                }
                else {
                    $('#order').hide();
                }
            });

            // Add a click event handler for the integrate button
            $('#integrate-button').click(doIntegration);

        });
    };

    /**
     * This method loads a table with the initial data and the graph.
     */
    function loadSampleData() {

        //Our sample data - Velocity vs Time
        var values = [
            [0, 0],
            [1, 1],
            [2, 4],
            [3, 8],
            [4, 14],
            [5, 21],
            [6, 28],
            [7, 35],
            [8, 43],
            [9, 51],
            [10, 58],
            [11, 64],
            [12, 69],
            [13, 73],
            [14, 75],
            [15, 73],
            [16, 68],
            [17, 60],
            [18, 49],
            [19, 35],
            [20, 33]
        ];

        Excel.run(function (ctx) {

            //Get the current active sheet
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();

            //Check if the table exist
            var table = sheet.tables.getItemOrNullObject("dataTable");

            return ctx.sync().then(function () {

                //If the table doesn't exist, make it, so it doesn't keep trying to over lap the same table every refresh
                if (table.isNullObject) {
                    //Create a table
                    table = sheet.tables.add('A1:B1', true);
                    table.name = "dataTable"; //Name of the table
                    table.getHeaderRowRange().values = [['x', 'y']];  //Column header names
                    table.rows.add(null, values); //Load dataset to the table

                    //Create Data headers
                    sheet.getRange("D18:D18").values = "Predicted Equation:";
                    sheet.getRange("D19:D19").values = "Approximate Integral:";
                    sheet.getRange("D19:D19").format.autofitColumns();

                    //Variable shortcut to give and access validation to the table
                    var tableValidation = table.getDataBodyRange().dataValidation;
                    tableValidation.ignoreBlanks = false; //false so our validation can also apply to blank cells

                    //Create a validation rule that only allow cells that contain numbers and are not blank
                    tableValidation.rule = {
                        custom: {
                            formula: "=AND(ISNUMBER(A2),NOT(ISBLANK(A2)))"
                        }
                    };

                    //Display error alert settings when user input invalid data
                    tableValidation.errorAlert = {
                        message: "Please enter a number",
                        showAlert: true,
                        style: "Information",
                        title: "Invalid Input"
                    };

                    //Add an event handler to the table, which will call checkForBlank on change
                    table.onChanged.add(checkForBlank);
                }

                return ctx.sync().then(function () {

                    //Generate a graph based on the loaded data
                    generateChart()
                }).then(ctx.sync);

            }).then(ctx.sync);
        }).catch(errorHandler);
    }

    /**
     * Create a chart based on the dataTable table
     */
    function generateChart() {
        Excel.run(function (ctx) {

            //Get the table data and its range
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var table = sheet.tables.getItem("dataTable");
            var bodyRange = table.getDataBodyRange();

            //Create a new scatter chart with the tables range and set it's name and position
            var chart = sheet.charts.add(Excel.ChartType.xyscatter, bodyRange, "Auto");
            chart.title.text = "Velocity vs. Time";
            chart.setPosition("D2");

            return ctx.sync().then(function () {

            }).then(ctx.sync);

        }).catch(errorHandler);
    }

    /**
     * Based on the users input, appropriately perform the correct regression and display the prediction
     * on the sheet
     * */
    function doRegression() {
        Excel.run(function (ctx) {
            //Grab the active worksheet, the data table, and load the values from the table
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var table = sheet.tables.getItem("dataTable");
            var bodyRange = table.getDataBodyRange().load("values");
            var isTableValid = table.getDataBodyRange().dataValidation.load("valid");

            return ctx.sync().then(function () {

                //using dataValidation.valid, we can check if our custom formula is still valid
                if (isTableValid.valid !== true) { //If it's not valid, we exit out of the function and diplay an error
                    $('#tableError').show();
                    sheet.getRange("E18:E18").values = "";
                    return;
                } else {
                    $('#tableError').hide();
                }

                //Get the table data and a container for the result of regression
                let xyData = bodyRange.values;
                var result = 0;

                //Create the regression based on the selected regression type
                if ($('#linear').is(':checked')) {
                    //Do linear regression
                    result = regression.linear(xyData);
                }
                else if ($('#exponential').is(':checked')) {
                    //Do exponential function regression
                    result = regression.exponential(xyData);
                }
                else if ($('#logarithmic').is(':checked')) {
                    //Do logarithmic regression
                    result = regression.logarithmic(xyData);
                }
                else if ($('#power').is(':checked')) {
                    //Do power function regression
                    result = regression.power(xyData);
                } else {
                    //Take the value in the order box (for polynomial regression)
                    //Use math.js evaluate to use as a variable
                    var ord = math.evaluate($('#orderBox').val());

                    //If the input we get is invalid, display error message
                    if (typeof ord === "undefined" || ord < 1 || ord > 10) {
                        $('#orderError').show();
                        sheet.getRange("E18:E18").values = "";
                        return;
                    } else {
                        $('#orderError').hide();
                    }
                    //Do polynomial regression with the order provided
                    result = regression.polynomial(xyData, { order: ord });
                }

                //Stringify the result
                var expression = result.string.slice(4);

                //Write the equation back to the sheet
                sheet.getRange("E18:E18").values = expression;

            }).then(ctx.sync);
        }).catch(errorHandler);
    }

    /**
     * Method to change the table depending any event changes in the table of dataTable.
     *  The event will listen to new, edited or removed cells and change their color depending if
     *  they are filled or not
     * @param {any} eventArgs
     */
    function checkForBlank(eventArgs) {
        Excel.run(function (ctx) {

            //Get the range that event occured on
            var range = ctx.workbook.worksheets.getActiveWorksheet().getRange(eventArgs.address);
            range.load("values, rowCount, columnCount")

            return ctx.sync().then(function () {
                //Loop through all the cells in the range
                for (var i = 0; i < range.rowCount; i++) {
                    for (var j = 0; j < range.columnCount; j++) {
                        //In event types of RowInserted or RangeEditied, we'll check the value in that cell
                        if (eventArgs.changeType === "RowInserted" || eventArgs.changeType === "RangeEdited") {
                            //If the cell is not a number, we'll highlight it red
                            if (typeof range.values[i][j] !== "number") {
                                range.getCell(i, j).format.fill.color = "red";
                            } else { //Else, we'll make sure there isn't any fill coloring in that cell, attempting to remove the red coloring
                                range.getCell(i, j).format.fill.clear();
                            }
                            //In a RowDeleted event, we'll clear any coloring and remove any values in that cell.
                        } else if (eventArgs.changeType === "RowDeleted") {
                            range.getCell(i, j).format.fill.clear();
                            range.getCell(i, j).values = ""
                        }
                    }
                }


            }).then(ctx.sync);

        }).catch(errorHandler);
    }

    /**
     * This method checks the user input and current state of the data set (for the x axis) 
     * before performing the integrate method. Any invalid inputs - handle the error appropriately
     */
    function doIntegration() {
        Excel.run(function (ctx) {

            //Gets data from the tables x column and the appropriate cell ranges.
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var xData = sheet.tables.getItem("dataTable").columns.getItem("x").getDataBodyRange().load("values");
            var expressionCell = sheet.getRange("E18:E18").load("values");
            var integrationCell = sheet.getRange("E19:E19").load("values");
            var isTableValid = sheet.tables.getItem("dataTable").getDataBodyRange().dataValidation.load("valid");
            

            return ctx.sync().then(function () {

                //Checking if our expressionCell is filled. otherwise, exit out of the function and handle error messages appropriately
                if (expressionCell.values[0][0] === "" || isTableValid.valid !== true) {
                    $('#integrateError').show();
                    integrationCell.values = "";

                    $('#lowerError').hide();
                    $('#upperError').hide();
                    //Precision will be implemented in a later version
                    //$('#precisionError').hide();
                    $('#boundsError').hide();
                    return;
                } else {
                    $('#integrateError').hide();
                }

                //Get boundary limits from the table
                var lowerX = xData.values[0];
                var upperX = xData.values[xData.values.length - 1];

                //get user input form the taskpane
                var lowerBound = Number($('#lowerBound').val());
                var upperBound = Number($('#upperBound').val());
                //Precision will be implemented in a later version
                //var precision = Number($('#precision').val());

                //Flag if we don't hit any of the if statements means all user input is valid
                var validIntegrationInput = true;

                //If states that check the field is not empty, it is not a number, and within bounds. It will handle the error messages appropriately
                if (!$('#lowerBound').val() || isNaN(lowerBound) || lowerBound < lowerX) {
                    $('#lowerError').show();
                    validIntegrationInput = false;
                } else {
                    $('#lowerError').hide();
                }

                if (!$('#upperBound').val() || isNaN(upperBound) || upperBound > upperX) {
                    $('#upperError').show();
                    validIntegrationInput = false;
                } else {
                    $('#upperError').hide();
                }

                //Precision will be implemented in a later version
                //if (!$('#precision').val() || isNaN(precision) || precision < 0.0001 || precision > 0.1) {
                //    $('#precisionError').show();
                //    validIntegrationInput = false;
                //} else {
                //    $('#precisionError').hide();
                //}

                //Error change if the user's upper and lower bounds don't cross each other, only if both inputs are actually given
                if ((lowerBound > upperBound) && (!isNaN(lowerBound) && !isNaN(upperBound)) && ($('#lowerBound').val() && $('#upperBound').val()) ) {
                    $('#boundsError').show();
                    validIntegrationInput = false;
                } else {
                    $('#boundsError').hide();
                }

                //If no error occured, call integrate method, otherwise handle error 
                if (validIntegrationInput) {
                    //Do the integrate method and display it on the sheet
                    //Precision will be implemented in a later version
                    integrationCell.values = integrate(expressionCell.values, upperBound, lowerBound, 1000);
                } else {
                    integrationCell.values = "";
                    return;
                }
                
            }).then(ctx.sync);
        }).catch(errorHandler);
    }

    /**
     * This function will employ Simpson's rule to integrate the expression returned from regression
     * @param {any} expression
     * @param {any} b
     * @param {any} a
     * @param {any} N
     */
    function integrate(expression, b, a, N) {
        var f = math.evaluate("f(x)=" + expression);

        //Use Simpson's rule to evaluate 
        var result = f(a) + f(b) + 4 * summarize(f, 1, N, halfStepper(b, a, N)) + 2 * summarize(f, 1, N - 1, xStepper(b, a, N));
        result = result * stepLength(b, a, N) / 6;

        return result;
    }

    //Calculate the length of each interval 
    function stepLength(b, a, N) {
        return (b - a) / N;
    }

    //This function will generate the functions describing each step along the length of the integral domain
    function xStepper(b, a, N) {
        var h = stepLength(b, a, N);
        return function (n) {
            return a + h * n;
        }
    }

    //Generates a midpoint function for each step along the integral domain
    function halfStepper(b, a, N) {
        var stepper = xStepper(b, a, N);
        return function (n) {
            return (stepper(n - 1) + stepper(n)) / 2;
        }
    }

    //Helper function to summarize the multiple functional values along the integral domain
    //where there are duplicate values
    function summarize(f, i, j, v) {
        var result = 0;
        for (var n = i; n < j + 1; n++) {
            result += f(v(n));
        }

        return result;
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();