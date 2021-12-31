/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

/*
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});
*/

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

//#region GLOBAL VARIABLES AND REFERENCES -------------------------------------------------------------------------

  //#region IMAGE REFERENCES --------------------------------------------------------------------------------------
  // import { ContextReplacementPlugin } from "webpack";
  import "../../assets/icon-16.png";
  import "../../assets/icon-32.png";
  import "../../assets/icon-80.png";
  //#endregion ----------------------------------------------------------------------------------------------------


  //#region GLOBAL VARIABLES --------------------------------------------------------------------------------------
  var artistColumn = "S";
  var moveEvent;
  var sortEvent;
  var sortColumn = "Priority";
  var projectTypeColumn = "H";
  var productColumn = "G";
  var addedColumn = "J";
  var loop = true;
  var startOverrideColumn = "U";
  var workOverrideColumn = "V";

  var brandNewBuild;
  var newBuildOtherNatives;
  var newBuildFromTemplate;
  var changesToExistingNatives;
  var specCheck;
  var weTransferUpload;
  var specialRequest;
  var otherProjectType;

  var menu;
  var menuXL;
  var smallMenu;
  var brochure;
  var brochureXL;
  var smallBrochure;
  var postcard;
  var jumboPostcard;
  var colossalPostcard;
  var scratchoffPostcard;
  var jumboScratchoffPostcard;
  var peelBoxPostcard;
  var magnet;
  var foldedMagnet;
  var twoSBT;
  var boxTopper;
  var flyer;
  var doorHanger;
  var smallPlastic;
  var mediumPlastic;
  var largePlastic;
  var couponBooklet;
  var envelopeMailer;
  var birthdayPostcard;
  var newMover;
  var plasticNewMover;
  var birthdayPlastic;
  var wideFormat;
  var windowClings;
  var businessCards;
  var artworkOnly;
  var logoCreation;
  var logoRecreation;
  var legalLetter;
  var letter;
  var mapCreation;
  var menuXXL;
  var biFoldMenu;
  var mediaKit;
  var popBanner;
  var otherProduct;

    //#region WEEKDAY VARIABLES ----------------------------------------------------------------------------------------------------------------

      var sunday = {
        dayID: 0,
        startHour: 8,
        startMinute: 30,
        endHour: 17,
        endMinute: 30,
        workDay: 0,
      }
      var monday = {
        dayID: 1,
        startHour: 8,
        startMinute: 0,
        endHour: 17,
        endMinute: 0,
        workDay: 0,
      }
      var tuesday = {
        dayID: 2,
        startHour: 8,
        startMinute: 30,
        endHour: 17,
        endMinute: 30,
        workDay: 0,
      }
      var wednesday = {
        dayID: 3,
        startHour: 8,
        startMinute: 30,
        endHour: 17,
        endMinute: 30,
        workDay: 0,
      }
      var thursday = {
        dayID: 4,
        startHour: 8,
        startMinute: 0,
        endHour: 18,
        endMinute: 0,
        workDay: 0,
      }
      var friday = {
        dayID: 5,
        startHour: 8,
        startMinute: 30,
        endHour: 13,
        endMinute: 30,
        workDay: 0,
      }
      var saturday = {
        dayID: 6,
        startHour: 8,
        startMinute: 30,
        endHour: 17,
        endMinute: 30,
        workDay: 0,
      }

      var weekdayList = [sunday, monday, tuesday, wednesday, thursday, friday, saturday];

    //#endregion --------------------------------------------------------------------------------------------------------------------------------

  //#endregion ----------------------------------------------------------------------------------------------

//#endregion ------------------------------------------------------------------------------------------------------



//#region BUTTON BEHAVIOR -----------------------------------------------------------------------------------------

  //#region CHECKBOX SETUP ________________________________________________________________________________________
  /** When the checkbox is CHANGED */
  $("#set-behavior").on("change", function() {

    // Is this set to checked?
    var checked = $(this).prop("checked");

    if (checked == true) { // Set the startup behavior!
      Office.addin.setStartupBehavior(Office.StartupBehavior.load); //when document opens, references startup behavioir in manifest, which automatically opens the taskpane
    } else { // Turn off the startup behavior!
      Office.addin.setStartupBehavior(Office.StartupBehavior.none); //when document opens, references startup behavioir in manifest, which automatically opens the taskpane
    }
  })
  //#endregion ----------------------------------------------------------------------------------------------------

//#endregion -------------------------------------------------------------------------------------------------------



//#region WATCHING FOR CHANGES ------------------------------------------------------------------------------------

  //#region STARTUP BEHAVIOR --------------------------------------------------------------------------------------
  Office.onReady((info) => {
    console.log("Office is ready!")
    // Load on Startup
    // setStartupBehavior is **document level**
    /*
    var currentBehavior = Office.addin.getStartupBehavior().then(function(returned) {
      if (returned == "Load") {
        //Check the checkbox
        $("#set-behavior").prop("checked", true);
      } else {
        //Uncheck the checkbox
        $("#set-behavior").prop("checked", false);
      }
      console.log(returned);
    });
    */
      if (info.host === Office.HostType.Excel) { //If application is Excel
        document.getElementById("sideload-msg").style.display = "none"; //Don't show side-loading message
        document.getElementById("app-body").style.display = "flex"; //Keep content in taskpane flexible to scaling, I think...
          
        Excel.run(async context => { //Do while Excel is running

          moveEvent = context.workbook.tables.onChanged.add(onTableChanged);

          // sortEvent = context.workbook.tables.onChanged.add(sortDate);

          return context.sync().then(function() { //Commits changes to document and then returns the console.log
            // console.log("Event handlers have been successfully registered");
          });
        });
      };
  });
  //#endregion ------------------------------------------------------------------------------------------------

//#endregion -----------------------------------------------------------------------------------------------------



//#region MOVING AND UPDATING DATA --------------------------------------------------------------------------------

  async function onTableChanged(eventArgs) { //This function will be using event arguments to collect data from the workbook
  // async function onTableChanged(eventArgs: Excel.TableChangedEventArgs) { //TypeScript version of this command

    await Excel.run(async (context) => {      

      //#region EVENT VARIABLES -----------------------------------------------------------------------------------

        var details = eventArgs.details; //Loads the values before and after the event
        var address = eventArgs.address; //Loads the cell's address where the event took place
        var changeType = eventArgs.changeType;
        var sheet = context.workbook.worksheets.getActiveWorksheet().load("name");
        var changedTable = context.workbook.tables.getItem(eventArgs.tableId).load("name"); //Returns tableId of the table where the event occured
        var regexStr = address.match(/[a-zA-Z]+|[0-9]+(?:\.[0-9]+|)/g); //Separates the column letter(s) from the row number for the address: presented as a string
        var changedColumn = regexStr[0]; //The first instance of the separated address array, being the column letter(s)
        var changedRow = Number(regexStr[1]) - 2; //The second instance of the separated address array, being the row, converted into a number and subtracted by 2
        var myRow = changedTable.rows.getItemAt(changedRow).load("values"); //loads the values of the changed row in the table where the event was fired 

        var addedAddress = "J" + (changedRow + 2); //takes the row that was updated and locates the address from the Added column.
        var addedRange = sheet.getRange(addedAddress);
        addedRange.load("values");

        var startAddress = "U" + (changedRow + 2);
        var startRange = sheet.getRange(startAddress);
        startRange.load("values");

        var workAddress = "V" + (changedRow + 2);
        var workRange = sheet.getRange(workAddress);
        workRange.load("values");

        var changedRowAddress = "A" + (changedRow + 2) + ":" + "V" + (changedRow + 2);
        var changedRange = sheet.getRange(changedRowAddress);
      
        //var unassignedTable = context.workbook.tables.getItem("UnassignedProjects");
        //var productColumn = changedTable.columns.getItem("Product").load("name");
        //var projectTypeColumn = changedTable.columns.getItem("Project Type").load("name");
        //var addedColumn = changedTable.columns.getItem("Added").load("name");
        //var artistColumn = changedTable.columns.getItem("Artist").load("name");

        var pickupTurnaroundTimeTable = context.workbook.tables.getItem("PickupTurnaroundTime");
        var projectTypeHourColumn = projectTypeTable.columns.getItem("Project Type Hour");
        projectTypeHourColumn.load("name");
        var pickupTurnaroundTimeTableRows = pickupTurnaroundTimeTable.rows
        pickupTurnaroundTimeTableRows.load("items")

        var productTable = context.workbook.tables.getItem("ProductTable");
        var productTableHoursColumn = productTable.columns.getItem("Product Hours");
        productTableHoursColumn.load("name");
        var productTableRows = productTable.rows
        productTableRows.load("items");

      //#endregion ------------------------------------------------------------------------------------------------


      //#region SPECIFIC TABLE VARIABLES --------------------------------------------------------------------------

        //#region UNASSIGNED PROJECTS VARIABLES ------------------------------------------------------------
          var unassignedSheet = context.workbook.worksheets.getItem("Unassigned Projects");
          var unassignedTable = unassignedSheet.tables.getItem("UnassignedProjects");
        //#endregion --------------------------------------------------------------------------

        //#region MATT VARIABLES --------------------------------------------------------
          var mattSheet = context.workbook.worksheets.getItem("Matt");
          var mattTable = mattSheet.tables.getItem("MattProjects");
        //#endregion --------------------------------------------------------------------------

        //#region ALAINA VARIABLES ------------------------------------------------------
          var alainaSheet = context.workbook.worksheets.getItem("Alaina");
          var alainaTable = alainaSheet.tables.getItem("AlainaProjects");
        //#endregion --------------------------------------------------------------------------

        //#region BERTO VARIABLES ------------------------------------------------------
          var bertoSheet = context.workbook.worksheets.getItem("Berto");
          var bertoTable = bertoSheet.tables.getItem("BertoProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region BRE B. VARIABLES ------------------------------------------------------
          var breBSheet = context.workbook.worksheets.getItem("Bre B.");
          var breBTable = breBSheet.tables.getItem("BreBProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region CHRISTIAN VARIABLES ------------------------------------------------------
          var christianSheet = context.workbook.worksheets.getItem("Christian");
          var christianTable = christianSheet.tables.getItem("ChristianProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region EMILY VARIABLES ------------------------------------------------------
          var emilySheet = context.workbook.worksheets.getItem("Emily");
          var emilyTable = emilySheet.tables.getItem("EmilyProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region IAN VARIABLES ------------------------------------------------------
          var ianSheet = context.workbook.worksheets.getItem("Ian");
          var ianTable = ianSheet.tables.getItem("IanProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region JEFF VARIABLES ------------------------------------------------------
          var jeffSheet = context.workbook.worksheets.getItem("Jeff");
          var jeffTable = jeffSheet.tables.getItem("JeffProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region JOSH VARIABLES ------------------------------------------------------
          var joshSheet = context.workbook.worksheets.getItem("Josh");
          var joshTable = joshSheet.tables.getItem("JoshProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region KRISTEN VARIABLES ------------------------------------------------------
          var kristenSheet = context.workbook.worksheets.getItem("Kristen");
          var kristenTable = kristenSheet.tables.getItem("KristenProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region ROBIN VARIABLES ------------------------------------------------------
          var robinSheet = context.workbook.worksheets.getItem("Robin");
          var robinTable = robinSheet.tables.getItem("RobinProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region LUKE VARIABLES ------------------------------------------------------
          var lukeSheet = context.workbook.worksheets.getItem("Luke");
          var lukeTable = lukeSheet.tables.getItem("LukeProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region LISA VARIABLES ------------------------------------------------------
          var lisaSheet = context.workbook.worksheets.getItem("Lisa");
          var lisaTable = lisaSheet.tables.getItem("LisaProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region LUIS VARIABLES ------------------------------------------------------
          var luisSheet = context.workbook.worksheets.getItem("Luis");
          var luisTable = luisSheet.tables.getItem("LuisProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region PETER VARIABLES ------------------------------------------------------
          var peterSheet = context.workbook.worksheets.getItem("Peter");
          var peterTable = peterSheet.tables.getItem("PeterProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region RITA VARIABLES ------------------------------------------------------
          var ritaSheet = context.workbook.worksheets.getItem("Rita");
          var ritaTable = ritaSheet.tables.getItem("RitaProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region ETHAN VARIABLES ------------------------------------------------------
          var ethanSheet = context.workbook.worksheets.getItem("Ethan");
          var ethanTable = ethanSheet.tables.getItem("EthanProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region BRE Z. VARIABLES ------------------------------------------------------
          var breZSheet = context.workbook.worksheets.getItem("Bre Z.");
          var breZTable = breZSheet.tables.getItem("BreZProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region JOE VARIABLES ------------------------------------------------------
          var joeSheet = context.workbook.worksheets.getItem("Joe");
          var joeTable = joeSheet.tables.getItem("JoeProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region JORDAN VARIABLES ------------------------------------------------------
          var jordanSheet = context.workbook.worksheets.getItem("Jordan");
          var jordanTable = jordanSheet.tables.getItem("JordanProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region HAZEL-RAH VARIABLES ------------------------------------------------------
          var hazelSheet = context.workbook.worksheets.getItem("Hazel-Rah");
          var hazelTable = hazelSheet.tables.getItem("HazelProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region TODD VARIABLES ------------------------------------------------------
          var toddSheet = context.workbook.worksheets.getItem("Todd");
          var toddTable = toddSheet.tables.getItem("ToddProjects");
        //#endregion ---------------------------------------------------------------------------

        //#region VALIDATION VARIABLES ------------------------------------------------------
          var validationSheet = context.workbook.worksheets.getItem("Validation");
          var toddTable = toddSheet.tables.getItem("ToddProjects");
        //#endregion ---------------------------------------------------------------------------

      //#endregion ------------------------------------------------------------------------------------------------


      //#region ON ROW INSERTED ----------------------------------------------------------------------------------- 

        if (changeType == "RowInserted") {

          //#region LOAD VARIABLES AND DO FUNCTIONS ---------------------------------------------------------------

            await context.sync().then(function () { //loads variable values

              var addedRangeValues = addedRange.values[0][0]; //loads cell values in the Added column
              var startRangeValues = startRange.values[0][0]; //loads cell values in the Start Override column
              var workRangeValues = workRange.values[0][0]; //loads cell values in the Work Override column

              //#region AUTOFILL ADDED COLUMN WITH CURRENT DATE/TIME ---------------------------------------------

                if (addedRangeValues == "") {
                  var newRange = currentDate(sheet, changedRow);
                  //return newRange;
                } else {
                console.log("Inserted row already had an Added date, so the current time was not assigned");
                };

              //#endregion ---------------------------------------------------------------------------------------

              //#region AUTOFILL OVERRIDE COLUMNS WITH 0 IF EMPTY ------------------------------------------------

                if (startRangeValues == "") {
                  startRangeValues = [["0"]];
                  //return startRangeValues;
                };

                if (workRangeValues == "") {
                  workRangeValues = [["0"]];
                  //return workRangeValues;
                };

              //#endregion ---------------------------------------------------------------------------------------

              //#region ERROR HANDLING -----------------------------------------------------------------------------

              }).catch(function (error) {
                console.log('Error: ' + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                };
                //console.log("Promise Rejected");
              });

            //#endregion -----------------------------------------------------------------------------------------

            };

          //#endregion -------------------------------------------------------------------------------------------------

      //#endregion --------------------------------------------------------------------------------------------------


      //#region ON RANGE EDITED ------------------------------------------------------------------------------------

        if (changeType == "RangeEdited" && eventArgs.details !== undefined ) {
          
          //#region ESCAPES ON TABLE CHANGED FUNCTION IF VALUES ARE UNCHANGED --------------------------------------

            // If values are the same as before, ignore the moved-to table's on change event        
            if (eventArgs.details.valueAfter == eventArgs.details.valueBefore) {
              //console.log("No values have changed. Exiting move data event...")
              return;
            };

          //#endregion --------------------------------------------------------------------------------------------
            
          //#region LOAD VARIABLES AND DO FUNCTIONS ----------------------------------------------------------------

            await context.sync().then(function () { //loads variable values
          
              //#region LOAD & ASSIGN TABLE VALUES --------------------------------------------------------------------

                var rowValues = myRow.values;

                console.log(pickupTurnaroundTimeTableRows.items);

                //if (changedTable.id == projectTypeTable.id) {
                  for (var i = 0; i < pickupTurnaroundTimeTableRows.items.length; i++) {
                    assignPickupTurnaroundTimeValues(pickupTurnaroundTimeTableRows, i); //loads value from validation sheet for each project type
                  };
                //};

                //if (changedTable.id == productTable.id) {
                  for (var i = 0; i < productTableRows.items.length; i++) {
                    assignProductValues(productTableRows, i); //loads value from validation sheet for each product
                  };
                //};

              //#endregion --------------------------------------------------------------------------------------------

              //#region CLEAN UP TEXT FORMATTING ----------------------------------------------------------------------

                changedRange.format.font.name = "Calibri";
                changedRange.format.font.size = 12;
                changedRange.format.font.color = "#000000";

              //#endregion --------------------------------------------------------------------------------------------

              //#region IF CHANGE WAS NOT MADE TO VALIDATION SHEET... ----------------------------------------------------

                if (sheet.id !== validationSheet.id) {

                  //#region ADJUSTING TURN AROUND TIME --------------------------------------------------------------------

                    if (changedColumn == projectTypeColumn || changedColumn == productColumn || changedColumn == addedColumn || changedColumn == startOverrideColumn || changedColumn == workOverrideColumn) { //if updated data was in Project Type column, run the lookupStart function

                      var projectTypeHours = lookupStart(rowValues, changedRow); //adds hours to turn-around time based on Project Type
                    
                      var productHours = preLookupWork(rowValues, projectTypeHours); //adds hours based on Product and adds to lookupStart output
                    
                      var workHoursAdjust = lookupWork(productHours, rowValues); //takes prelookupWork variable and divides by 3 if lookupStart was equal to 2. Otherwise remains the same.
                
                      var myDate = receivedAdjust(rowValues, changedRow); //grabs values from Added column and converts into date object in EST.
                    
                      var override = startPreAdjust(rowValues, projectTypeHours, myDate); //adds manual override start hours to adjusted start time. Adjusts for office hours and weekends.
                    
                      var startedPickedUpBy = startedBy(changedRow, sheet, override); //Prints the value of override to the Picked Up / Started By column and formats the date in a readible format.
                
                      var workOverride = workPrePreAdjust(rowValues, workHoursAdjust, override); //Finds the value of Work Override in the changed row and adds it to workHoursAdjust, then adds that new number as hours to startedPickedUpBy. Formats to be within office hours and on a weekday if needed.
                  
                      var proofToClient = toClient(changedRow, sheet, workOverride); //Prints the value of workOverride to the Proof to Client column and formats the date in a readible format.

                      console.log("Turn Around time variables were updated!");

                      return;
                  
                  };

                  //#endregion ------------------------------------------------------------------------------------------


                  //#region MOVE DATA BETWEEN TABLES ---------------------------------------------------------------------

                    if (changedColumn == artistColumn) { //if updated data was in the Artist column, run the following code

                      if (details.valueAfter == "Unassigned") {
                        unassignedTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Unassigned Projects Table!");
                        return;
                      } else if (details.valueAfter == "Matt") {
                        mattTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Matt Table!");
                        return;
                      } else if (details.valueAfter == "Alaina") {
                        alainaTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Alaina Table!");
                        return;            
                      } else if (details.valueAfter == "Berto") {
                        bertoTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Berto Table!");
                        return;
                      } else if (details.valueAfter == "Bre B.") {
                        breBTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Bre B. Table!");
                        return;
                      } else if (details.valueAfter == "Christian") {
                        christianTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Christian Table!");
                        return;
                      } else if (details.valueAfter == "Emily") {
                        emilyTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Emily Table!");
                        return;
                      } else if (details.valueAfter == "Ian") {
                        ianTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Ian Table!");
                        return;
                      } else if (details.valueAfter == "Jeff") {
                        jeffTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Jeff Table!");
                        return;
                      } else if (details.valueAfter == "Josh") {
                        joshTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Josh Table!");
                        return;
                      } else if (details.valueAfter == "Kristen") {
                        kristenTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Kristen Table!");
                        return;
                      } else if (details.valueAfter == "Robin") {
                        robinTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Robin Table!");
                        return;
                      } else if (details.valueAfter == "Luke") {
                        lukeTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Luke Table!");
                        return;
                      } else if (details.valueAfter == "Lisa") {
                        lisaTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Lisa Table!");
                        return;
                      } else if (details.valueAfter == "Luis") {
                        luisTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Luis Table!");
                        return;
                      } else if (details.valueAfter == "Peter") {
                        peterTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Peter Table!");
                        return;
                      } else if (details.valueAfter == "Rita") {
                        ritaTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Rita Table!");
                        return;
                      } else if (details.valueAfter == "Ethan") {
                        ethanTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Ethan Table!");
                        return;
                      } else if (details.valueAfter == "Bre Z.") {
                        breZTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Bre Z. Table!");
                        return;
                      } else if (details.valueAfter == "Joe") {
                        joeTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Joe Table!");
                        return;
                      } else if (details.valueAfter == "Jordan") {
                        jordanTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Jordan Table!");
                        return;
                      } else if (details.valueAfter == "Hazel-Rah") {
                        hazelTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Hazel-Rah Table!");
                        return;
                      } else if (details.valueAfter == "Todd") {
                        toddTable.rows.add(null, myRow.values); //Adds empty row to bottom of GreenBasket Table, then inserts the changed values into this empty row
                        myRow.delete(); //Deletes the changed row from the original sheet
                        console.log("Data was moved to the Todd Table!");
                        return;
                      } else {
                        console.log("Looks like there wasn't an Artist change this time. No data was moved...");
                      } return;
                    } else {
                      //console.log("The artist column was not updated, so nothing was moved...");
                      return;
                    };

                  //#endregion ----------------------------------------------------------------------------------------------------
                  
                } else {
                  console.log("Adjustments were made to the validation sheet, therefore the date variables and move functions were not triggered");
                };
                // context.sync();

              //#endregion ------------------------------------------------------------------------------------------------

              //#region ERROR HANDLING -------------------------------------------------------------------------------------

                }).catch(function (error) {
                  console.log('Error: ' + error);
                  if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                  }
                  //console.log("Promise Rejected");
                });

              //#endregion ------------------------------------------------------------------------------------------------

          //#endregion -----------------------------------------------------------------------------------------------

        }; 

      //#endregion ---------------------------------------------------------------------------------------------------

    });
  };

//#endregion ------------------------------------------------------------------------------------------------------



//#region SORTING -------------------------------------------------------------------------------------------------

  //#region SORT BY DATE ------------------------------------------------------------------------------------------
    async function sortDate(eventArgs) { //This function will be using event arguments to collect data from the workbook
      // console.log("SORT FUNCTION FIRED!");
      // console.log(eventArgs);

      var theChange = eventArgs.changeType; //Kind of change that was made
      var theDetails = eventArgs.details;

      // console.log("args ");

      
      if (theChange == "RangeEdited" && (theDetails == undefined || theDetails.valueTypeAfter == "String")) { //&& theDetails == undefined) {
        console.log("The sorting event has been initiated!!"); //Prevents an event from being triggered when a new row is inserted into the other sheet, thus causing duplicate runs

        //#region SORTING VARIABLES ---------------------------------------------------------------------------------
        Excel.run(async context => {
          var changedTable = context.workbook.tables.getItem(eventArgs.tableId); //Returns tableId of the table where the event occured
          var tableRange = changedTable.getRange(); //Gets the range of the changed table
          var sortHeader = tableRange.find(sortColumn, {}); //Gets the range of the entire sortColumn (the "Date" column) from the changed table
          sortHeader.load("columnIndex");
          sortHeader.load("addressLocal")
          // var sortTag = ["Urgent", "Semi-Urgent", "Not Urgent", "Eventual", "Downtime"];
          // const list = [
          //   { Tag: 'Urgent'},
          //   { Tag: 'Semi-Urgent'},
          //   { Tag: 'Not Urgent'},
          //   { Tag: 'Eventual'},
          //   { Tag: 'Downtime'},
          // ]
          //#endregion --------------------------------------------------------------------------------------------------

        //#region SORTING CONDITIONS --------------------------------------------------------------------------------
          return context.sync().then(function() {
            console.log("Sync completed...Ready to sort")
            // console.log(sortHeader.addressLocal);
            // console.log(list);

            // if (sortHeader.columnIndex == 14) {
            //   list.sort((a, b) => (a.Tag < b.Tag) ? 1 : -1);
            //   console.log(list);
            // }

            tableRange.sort.apply(
              [
                { //list of conditions to sort on
                  key: sortHeader.columnIndex, //sorts based on data in Date column
                  sortOn: Excel.SortOn.value, //sorts based on cell vlaues
                  ascending: true
                  // subField: Excel.subField, //sorts based on cell vlaues
                  // subField: String(sortTag)
                }
              ],
              false, //will not impact string ordering
              true, //table has headers
              Excel.SortOrientation.rows //sorts the rows based on previous conditions
            );

            // const myArray = [1, 2, 3, 4, 5, 6];
            // let filteredArray = list.filter((x) => {
            //   return x % 2 === 0;
            // });
            

        

            // Queue a command to apply a filter on the Category column
            // var filter = changedTable.columns.getItem("Tags").filter;
            // filter.apply({
            //     filterOn: Excel.FilterOn.values,
            //     values: ["Urgent", "Semi-Urgent", "Not Urgent", "Eventual", "Downtime"]
            // });



            console.log("Sorting is completed.")
          }); 
          //#endregion --------------------------------------------------------------------------------------------------

        }).catch(tryCatch); // CATCH EXCEL.RUN
      
      }; // END IF  
    } // END SORTDATE()
  //#endregion ----------------------------------------------------------------------------------------------------

//#endregion ------------------------------------------------------------------------------------------------------



//#region AUTOFILL FUNCTIONS -------------------------------------------------------------------------------------


  //#region CURRENT DATE & TIME IN ADDED COLUMN -------------------------------------------------------------------

    /**
     * Inputs the current date & time into the Added column of the changed row
     * @param {Object} sheet The active worksheet
     * @param {Number} changedRow The number of the changed row
     * @returns Array
     */
    function currentDate(sheet, changedRow) {

      var address = "J" + (changedRow + 2); //takes the row that was updated and locates the address from the Added column.
      var range = sheet.getRange(address);

      var now = new Date();
      var toSerial = JSDateToExcelDate(now);

      range.values = [[toSerial]];
      return range.values;

    };

  //#endregion ----------------------------------------------------------------------------------------------------


  //#region ASSIGN PROJECT TYPE VALUES FROM VALIDATION SHEET TO CODED VARIABLES -------------------------------------
    /**
     * Loads values from the Picked-Up / Started By Turn Around Time Table in Validation and assigns these values to the associated variable to be used in the code
     * @param {Array} pickupTurnaroundTimeTableRows An array of all the rows in the Picked-Up / Started By Turn Around Time table
     * @param {Number} i The number of the row that values will be assigned to
     */
    function assignPickupTurnaroundTimeValues(pickupTurnaroundTimeTableRows, i) {
      var pickupTurnaroundTimeValues = pickupTurnaroundTimeTableRows.items[i].values;
        //console.log(pickupTurnaroundTimeValues);
          if (i == 0) {
            brandNewBuild = pickupTurnaroundTimeValues[0][1];
          } else if (i == 1) {
            newBuildOtherNatives = pickupTurnaroundTimeValues[0][1];
          } else if (i == 2) {
            newBuildFromTemplate = pickupTurnaroundTimeValues[0][1];
          } else if (i == 3) {
            changesToExistingNatives = pickupTurnaroundTimeValues[0][1];
          } else if (i == 4) {
            specCheck = pickupTurnaroundTimeValues[0][1];
          } else if (i == 5) {
            weTransferUpload = pickupTurnaroundTimeValues[0][1];
          } else if (i == 6) {
            specialRequest = pickupTurnaroundTimeValues[0][1];
          } else if (i == 7) {
            otherProjectType = pickupTurnaroundTimeValues[0][1];
          };
    };
      //console.log("The updated projectType values are: " + brandNewBuild + ", " + newBuildOtherNatives + ", " + newBuildFromTemplate + ", " + changesToExistingNatives + ", " + specCheck + ", " + weTransferUpload + ", " + specialRequest + ", " + otherProjectType)

  //#endregion --------------------------------------------------------------------------------------------------------
        

  //#region ASSIGN PRODUCT VALUES FROM VALIDATION SHEET TO CODED VARIABLES --------------------------------------------
    /**
     * Loads values from the Product Table in Validation and assigns these values to the associated variable to be used in the code
     * @param {Array} productTableRows An array of all the rows in the Product table
     * @param {Number*} i The number of the row that values will be assigned to
     */
    function assignProductValues(productTableRows, i) {     
    var productValues = productTableRows.items[i].values;
      if (i == 0) {
        menu = productValues[0][1];
      } else if (i == 1) {
        menuXL = productValues[0][1];
      } else if (i == 2) {
        smallMenu = productValues[0][1];
      } else if (i == 3) {
        brochure = productValues[0][1];
      } else if (i == 4) {
        brochureXL = productValues[0][1];
      } else if (i == 5) {
        smallBrochure = productValues[0][1];
      } else if (i == 6) {
        postcard = productValues[0][1];
      } else if (i == 7) {
        jumboPostcard = productValues[0][1];
      } else if (i == 8) {
        colossalPostcard = productValues[0][1];
      } else if (i == 9) {
        scratchoffPostcard = productValues[0][1];
      } else if (i == 10) {
        jumboScratchoffPostcard = productValues[0][1];
      } else if (i == 11) {
        peelBoxPostcard = productValues[0][1];
      } else if (i == 12) {
        magnet = productValues[0][1];
      } else if (i == 13) {
        foldedMagnet = productValues[0][1];
      } else if (i == 14) {
        twoSBT = productValues[0][1];
      } else if (i == 15) {
        boxTopper = productValues[0][1];
      } else if (i == 16) {
        flyer = productValues[0][1];
      } else if (i == 17) {
        doorHanger = productValues[0][1];
      } else if (i == 18) {
        smallPlastic = productValues[0][1];
      } else if (i == 19) {
        mediumPlastic = productValues[0][1];
      } else if (i == 20) {
        largePlastic = productValues[0][1];
      } else if (i == 21) {
        couponBooklet = productValues[0][1];
      } else if (i == 22) {
        envelopeMailer = productValues[0][1];
      } else if (i == 23) {
        birthdayPostcard = productValues[0][1];
      } else if (i == 24) {
        newMover = productValues[0][1];
      } else if (i == 25) {
        plasticNewMover = productValues[0][1];
      } else if (i == 26) {
        birthdayPlastic = productValues[0][1];
      } else if (i == 27) {
        wideFormat = productValues[0][1];
      } else if (i == 28) {
        windowClings = productValues[0][1];
      } else if (i == 29) {
        businessCards = productValues[0][1];
      } else if (i == 30) {
        artworkOnly = productValues[0][1];
      } else if (i == 31) {
        logoCreation = productValues[0][1];
      } else if (i == 32) {
        logoRecreation = productValues[0][1];
      } else if (i == 33) {
        legalLetter = productValues[0][1];
      } else if (i == 34) {
        letter = productValues[0][1];
      } else if (i == 35) {
        mapCreation = productValues[0][1];
      } else if (i == 36) {
        menuXXL = productValues[0][1];
      } else if (i == 37) {
        biFoldMenu = productValues[0][1];
      } else if (i == 38) {
        mediaKit = productValues[0][1];
      } else if (i == 39) {
        popBanner = productValues[0][1];
      } else if (i == 40) {
        otherProduct = productValues[0][1];
      };
  };
  //console.log("The updated data for 2SBT is: " + twoSBT);
  //console.log("The updated date for Other is: " + otherProduct);

//#endregion ----------------------------------------------------------------------------------------------------


  //#region PROJECT TYPE HOURS -----------------------------------------------------------------------------------
    /**
     * Finds the value of Project Type in the changed row and returns a number of hours depending on the project type
     * @param {Array} rowValues loads the values of the changed row
     * @param {Number} changedRow loads the row number of the changed row
     * @returns Number
     */   
    function lookupStart(rowValues, changedRow) { //loads these variables from another function to use in this function
      var address = "H" + (changedRow + 2); //takes the row that was updated and locates the address from the Project Type column.
      // console.log("The address of the new Project Type is " + address);
      var input = rowValues[0][7]; //assigns input the cell value in the changed row and the Project Type column (a nested array of values)
      // console.log(input);

      var output;

      if (input == "Brand New Build") {
        output = brandNewBuild;
      } else if (input == "Brand New Build from Other Product Natives") {
        output = newBuildOtherNatives;
      } else if (input == "Brand New Build From Template") {
        output = newBuildFromTemplate;
      } else if (input == "Changes to Exisiting Natives") {
        output = changesToExistingNatives;
      } else if (input == "Specification Check") {
        output = specCheck;
      } else if (input == "WeTransfer Upload to MS") {
        output = weTransferUpload;
      } else if (input == "Special Request") {
        output = specialRequest;
      } else {
        output = otherProjectType;
      };

      return output;

    };
  
  //#endregion ---------------------------------------------------------------------------------------------------


  //#region PICKED UP / STARTED BY -------------------------------------------------------------------------------

    //References the Project Type column (H), Added column (J), and the Start Override column (U) to return a specific date and time for the project to by picked up by. This value is returned in the Picked Up / Started By column (M).

    //#region MY DATE ----------------------------------------------------------------------------------------------
      /**
       * Finds the value of Date Added in the changed row and converts it to be a date object in EST.
       * @param rowValues loads the values of the changed row
       * @param changedRow loads the row number of the changed row
       * @returns Date
       */
      function receivedAdjust(rowValues, changedRow) {
        var address = "J" + (changedRow + 2); //takes the row that was updated and locates the address from the Added column.
        var dateTime = rowValues[0][9]; //assigns input the cell value in the changed row and the Added column (a nested array of values)

        var date = new Date(Math.round((dateTime - 25569)*86400*1000)); //convert serial number to date object
        date.setMinutes(date.getMinutes() + date.getTimezoneOffset()); //adjusting from GMT to EST (adds 4 hours)
        return date;
      };
    //#endregion ---------------------------------------------------------------------------------------------------

    //#region OVERRIDE ---------------------------------------------------------------------------------------------
      /**
       * Finds the value of Start Override in the changed row and adds it to projectTypeHours, then adds that new number as hours to myDate. Adjusts for office hours and weekends.
       * @param {Array} rowValues loads the values of the changed row
       * @param {Number} projectTypeHours lookupStart returned number
       * @param {Date} myDate receivedAdjust returned date
       * @return {Date}
       */
      function startPreAdjust(rowValues, projectTypeHours, myDate) {
        var startOverride = rowValues[0][20]; //gets values of Start Orverride cell
        var startManualOverride = projectTypeHours + startOverride; //adds start override value to the number of hours for the project type
        var myDateCopy = new Date(myDate); //sets myDateCopy to myDate as a new date variable (so the old date doesnt get changed)
        var adjustedDateTime = officeHours(myDateCopy, startManualOverride); //converts to be within office hours if it already isn't
        return adjustedDateTime;
      };

    //#endregion ----------------------------------------------------------------------------------------------------

    //#region STARTED PICKED UP BY ---------------------------------------------------------------------------------
      /**
       * Prints the value of override to the Picked Up / Started By column and formats the date in a readible format
       * @param {Number} changedRow loads the row number of the changed row
       * @param {Object} sheet the active worksheet
       * @param {Date} override date adjusted for office hours
       * @returns date
       */
      function startedBy(changedRow, sheet, override) { //loads these variables from another function to use in this function
        var address = "M" + (changedRow + 2); //takes the row that was updated and locates the address from the Picked Up / Started By column.
        var range = sheet.getRange(address); //assigns the cell from the address variable to range
      
        /*
        the region below sets a custom cell format for the date so that it is more easily readible. It is not currently being used 
        because we decided later to apply some conditional formatting to the date cells, but excel didn't recognize our custom format as a date;
        instead I decided to convert the date object back into a serial number and then format the column in Excel to achieve the desired output
        */

        //#region FORMATTING DATE INTO READIBLE STRING ---------------------------------------------------------------
          /*
          var formatDate = override.toLocaleDateString("en-us", { //formats the date to display correctly
              weekday:'short',
              month:'numeric',
              day: 'numeric',
              year: '2-digit'
          });

          var formatTime = override.toLocaleTimeString("en-us", { //formats the time to display correctly
            hour: '2-digit',
            minute:'2-digit'
          });

          var squeekday = formatDate + " " + formatTime; //adds the correctly displayed date and time together

          range.values = [[squeekday]]; //assigns the returned date value to the cell

          return range.values; //commits changes and exits the function
          */
        //#endregion ------------------------------------------------------------------------------------------------
      
        var serialDate = JSDateToExcelDate(override);  //converts override date object back into a serial number

        range.values = [[serialDate]]; //assigns the returned serial number to the cell
        return range.values; //commits changes and exits the function

      };
    //#endregion ----------------------------------------------------------------------------------------------------

  //#endregion ------------------------------------------------------------------------------------------------------


  //#region PROOF TO CLIENT --------------------------------------------------------------------------------------

    //References the Project Type column (H), Product column (G), and the Work Override column (V) to return a specific date and time for a proof to be sent to the client. This value is returned in the Proof to Client column (N).

    //#region PRODUCT HOURS ----------------------------------------------------------------------------------------
      /**
       * Finds the value of Product in the changed row, returns a number of hours based on the product, and adds this number to projectTypeHours
       * @param {Array} rowValues loads the values of the changed row
       * @param {Number} projectTypeHours lookupStart returned number
       * @returns A Number
       */
      function preLookupWork(rowValues, projectTypeHours) {
        var input = rowValues[0][6]; //assigns input the cell value in the changed row and the Product column (a nested array of values)
        // console.log(input);
        var output;

        if (input == "Menu") {
          output = menu;
        } else if (input == "MenuXL") {
          output = menuXL;
        } else if (input == "Small Menu") {
          output = smallMenu;
        } else if (input == "Brochure") {
          output = brochure;
        } else if (input == "BrochureXL") {
          output = brochureXL;
        } else if (input == "Small Brochure") {
          output = smallBrochure;
        } else if (input == "Postcard") {
          output = postcard;
        } else if (input == "Jumbo Postcard") {
          output = jumboPostcard;
        } else if (input == "Colossal Postcard") {
          output = colossalPostcard;
        } else if (input == "Scratch-Off Postcard") {
          output = scratchoffPostcard;
        } else if (input == "Jumbo Scratch-Off Postcard") {
          output = jumboScratchoffPostcard;
        } else if (input == "Peel-A-Box Postcard") {
          output = peelBoxPostcard;
        } else if (input == "Magnet") {
          output = magnet;
        } else if (input == "Folded Magnet") {
          output = foldedMagnet;
        } else if (input == "2SBT") {
          output = twoSBT;
        } else if (input == "Box Topper") {
          output = boxTopper;
        } else if (input == "Flyer") {
          output = flyer;
        } else if (input == "Door Hanger") {
          output = doorHanger;
        } else if (input == "Small Plastic") {
          output = smallPlastic;
        } else if (input == "Medium Plastic") {
          output = mediumPlastic;
        } else if (input == "Large Plastic") {
          output = largePlastic;
        } else if (input == "Coupon Booklet") {
          output = couponBooklet;
        } else if (input == "Envelope Mailer") {
          output = envelopeMailer;
        } else if (input == "Birthday Postcard") {
          output = birthdayPostcard;
        } else if (input == "New Mover") {
          output = newMover;
        } else if (input == "Plastic New Mover") {
          output = plasticNewMover;
        } else if (input == "Birthday Plastic") {
          output = birthdayPlastic;
        } else if (input == "Wide Format") {
          output = wideFormat;
        } else if (input == "Window Clings") {
          output = windowClings;
        } else if (input == "Business Cards") {
          output = businessCards;
        } else if (input == "Artwork Only") {
          output = artworkOnly;
        } else if (input == "Logo Creation") {
          output = logoCreation;
        } else if (input == "Logo Recreation") {
          output = logoRecreation;
        } else if (input == "Legal Letter") {
          output = legalLetter;
        } else if (input == "Letter") {
          output = letter;
        } else if (input == "Map Creation") {
          output = mapCreation;
        } else if (input == "MenuXXL") {
          output = menuXXL;
        } else if (input == "Bi-Fold Menu") {
          output = biFoldMenu;
        } else if (input == "Media Kit") {
          output = mediaKit;
        } else if (input == "POP Banner") {
          output = popBanner;
        } else {
          output = otherProduct;
        };

        var newOutput = output + projectTypeHours; //adds hours from lookupStart to output and assigns new output to global variable
        // console.log(newOutput);
        return newOutput;

      };

    //#endregion --------------------------------------------------------------------------------------------------

    //#region WORK HOURS ADJUST ------------------------------------------------------------------------------------
      /**
       * if Project Type value is anything other than a new build (and friends), adjusts the Product Hours number to be a third of it's normal value, resulting in a shorter proof to client time
       * @param {Number} productHours preLookupWork returned number
       * @param {Number} rowValues loads the values of the changed row
       * @returns Number
       */
      function lookupWork(productHours, rowValues) {

        var input = rowValues[0][7]; //assigns input the cell value in the changed row and the Project Type column (a nested array of values)

        if(input == "Brand New Build" || input == "Special Request" || input == "Other") { //if input from Project Type column is any of these values...
          return productHours; //returns the productHours number unaltered
        };
        var output = productHours / 3; //returns the productHours number divided by 3
        return output;
      };
    //#endregion ---------------------------------------------------------------------------------------------------

    //#region WORKOVERRIDE --------------------------------------------------------------------------------------------
      /**
       * Finds the value of Work Override in the changed row and adds it to workHoursAdjust, then adds that new number as hours to startedPickedUpBy. Formats to be within office hours and on a weekday if needed.
       * @param {Array} rowValues loads the values of the changed row
       * @param {Number} workHoursAdjust loads the values of lookupWork
       * @param {Date} startedPickedUpBy loads the date that the project should be picked up by
       * @returns Date
       */
      function workPrePreAdjust (rowValues, workHoursAdjust, override) {
        var workOverride = rowValues[0][21]; //gets values of Work Orverride cell
        var workManualAdjust = workHoursAdjust + workOverride; //adds start override value to the number of hours for the project type
        var overrideCopy = new Date(override); //sets overrideCopy to a new date variable (so the old date doesnt get changed)
        var adjustedDateTime = officeHours(overrideCopy, workManualAdjust);
        return adjustedDateTime;
      };
    //#endregion --------------------------------------------------------------------------------------------------

    //#region PROOF TO CLIENT ---------------------------------------------------------------------------------
      /**
       * Prints the value of workOverride to the Proof to Client column and formats the date in a readible format
       * @param {Number} changedRow loads the row number of the changed row
       * @param {Object} sheet the active worksheet
       * @param {Date} workOverride proof to client date found in the workPreAdjust function (after converted to be within office hours and on a weekday)
       * @returns date
       */
      function toClient(changedRow, sheet, workOverride) { //loads these variables from another function to use in this function
        var address = "N" + (changedRow + 2); //takes the row that was updated and locates the address from the Proof to Client column.
        var range = sheet.getRange(address); //assigns the cell from the address variable to range

        /*
        the region below sets a custom cell format for the date so that it is more easily readible. It is not currently being used 
        because we decided later to apply some conditional formatting to the date cells, but excel didn't recognize our custom format as a date;
        instead I decided to convert the date object back into a serial number and then format the column in Excel to achieve the desired output
        */

        //#region FORMATTING DATE INTO READIBLE STRING ---------------------------------------------------------------
          /*
          var formatDate = workOverride.toLocaleDateString("en-us", { //formats the date to display correctly
              weekday:'short',
              month:'numeric',
              day: 'numeric',
              year: '2-digit'
          });

          var formatTime = workOverride.toLocaleTimeString("en-us", { //formats the time to display correctly
            hour: '2-digit',
            minute:'2-digit'
          });

          var squeekday = formatDate + " " + formatTime; //adds the correctly displayed date and time together

          range.values = [[squeekday]]; //assigns the returned date value to the cell

          return range.values; //commits changes and exits the function
          */
      //#endregion -------------------------------------------------------------------------------------------------

        var serialDateTheSecond = JSDateToExcelDate(workOverride); //converts workOverride date object back into a serial number

        range.values = [[serialDateTheSecond]]; //assigns the returned serial number to the cell
        return range.values; //commits changes and exits the function

      };
    //#endregion ----------------------------------------------------------------------------------------------------

  //#endregion ------------------------------------------------------------------------------------------------------


  //#region OFFICE HOURS ---------------------------------------------------------------------------------------
    /**
     * Sets weekday variables and loops through the withinOfficeHours function, which adjusts the date to be within office hours
     * @param {Date} date Date to be adjusted to be within office hours
     * @param {Number} number Number of adjustment hours to add to date
     * @returns Date
     */
    function officeHours(day, number) {

      //#region SETTING WORKDAY HOURS IN THE WEEKDAY VARIABLES -------------------------------------------------------------------------------------

        //loops through my weekday variables, finds returns the proper variable title for it's index in the array, and then runs it through the findWorkDay function
        for (var i = 0; i < weekdayList.length; i++) {
          var weekdayReplacement = findWorkDay(weekdayList[i]);
        };

      //#endregion --------------------------------------------------------------------------------------------------------------------------------

      //var aNum = 0

      while (loop == true) {
      var officeHours = withinOfficeHours(day, number);
      day = officeHours.date;
      number = officeHours.adjustmentNumber;
      loop = officeHours.loop;
      //aNum++
      };
      //console.log("The correct date & time is: " + day);
      loop = true;
      return day;
    };

      //#region FUNCTIONS -------------------------------------------------------------------------------------------------------------------------

        //#region WITHIN OFFICE HOURS FUNCTION -------------------------------------------------------------------------------------------------
          /**
           * Adjusts date to be within office hours while maintaining an accurate turn around time variable for the adjustment number
           * @param {Date} date Date to be adjusted to be within office hours 
           * @param {Number} adjustmentNumber Number of adjustment hours to add to date
           * @returns An object with properties (date, adjustment number, and loop)
           */
          function withinOfficeHours(date, adjustmentNumber) {

            //#region VARIABLES ------------------------------------------------------------------------------------------------------------

              //#region SETS DATE VARIABLES ----------------------------------------------------------------------------------------------

                //converts our input variables into milliseconds
                var dateMilli = date.getTime();
                var adjustmentNumberMilli = adjustmentNumber * 3600000;

                //gets day of the week attributes for the date variable
                var dateDayOfWeek = dayOfWeek(date); //returns a dayID (0-6) for the day of the week of the date object
                var dayTitle = titleDOW(dateDayOfWeek); //returns a day title based on the dayID of the dateDayOfWeek variable

                //retrives workday variables associated with the weekday of the date variable
                var bookendVars = startEndMidnight(date, dayTitle);

                  //#region ADJUSTS DATES IN CASE REQUEST WAS SUBMITTED OUTSIDE OF OFFICE HOURS ---------------------------------------

                    if (date < bookendVars.startOfWorkDayMilli) { //if date is between 12AM and start time, adjust hours to be the start time
                        date.setHours(dayTitle.startHour);
                        date.setMinutes(dayTitle.startMinute);
                        date.setSeconds(0);
                        dateMilli = date.getTime();
                        bookendVars = startEndMidnight(date, dayTitle);
                    };

                    if (date > bookendVars.endOfWorkDayMilli) { //if date is after end time and before 12AM, go to next day and adjust hours to be the start time of that next day
                        date.setDate(date.getDate() + 1);
                        dateDayOfWeek = dayOfWeek(date);
                        dayTitle = titleDOW(dateDayOfWeek);
                        date.setHours(dayTitle.startHour);
                        date.setMinutes(dayTitle.startMinute);
                        date.setSeconds(0);
                        dateMilli = date.getTime();
                        bookendVars = startEndMidnight(date, dayTitle);
                    };
                  
                  //#endregion ------------------------------------------------------------------------------------------------------------

                  //#region ADJUSTS DATES IN CASE REQUEST WAS SUBMITTED ON WEEKEND ----------------------------------------------------

                        if ((dateDayOfWeek == 6) || (dateDayOfWeek == 0)) { //if date was submitted on a weekend...
                          date = weekendAdjust(date, dateDayOfWeek);
                          dateDayOfWeek = dayOfWeek(date);
                          dayTitle = titleDOW(dateDayOfWeek);
                          date.setHours(dayTitle.startHour);
                          date.setMinutes(dayTitle.startMinute);
                          date.setSeconds(0);
                          dateMilli = date.getTime();
                          bookendVars = startEndMidnight(date, dayTitle);
                        };
              
                      //#endregion ------------------------------------------------------------------------------------------------------------

              //#endregion ----------------------------------------------------------------------------------------------------------------

              //#region SETS ADJUSTMENT DATE VARIABLES -----------------------------------------------------------------------------------

                //adds adjustmentNumber to date to get an adjustedDate value that will be used in later checks and calculations
                var adjustedDate = new Date(date);
                var adjustedDateMilli = adjustedDate.getTime();
                adjustedDateMilli = adjustedDateMilli + adjustmentNumberMilli;
                adjustedDate = new Date(adjustedDateMilli);

              //#endregion ---------------------------------------------------------------------------------------------------------------

              //#region SETS ADD A DAY VARIABLES -----------------------------------------------------------------------------------------

                  //gets day of the week attributes for the day after the date variable
                    var nextDay = new Date(date);

                    var newNextDay = getNextDay(nextDay); //also sets this variable to the start time of the next day
                    var addADay = newNextDay.nextDay;
                    var addADayTitle = newNextDay.nextDayTitle;
                    var addADayMilli = addADay.getTime();
                    
                    //retrives workday variables associated with the weekday of the addADay variable
                    var bookendAddedDate = startEndMidnight(addADay, addADayTitle);

                //#endregion ----------------------------------------------------------------------------------------------------------------

            //#endregion --------------------------------------------------------------------------------------------------------------------

            //#region ACTION: SETS ADJUSTED DATE TO BE WITHIN OFFICE HOURS ------------------------------------------------------------------

              //if adjustedDate falls outside of office hours, do this...
              if (adjustedDateMilli < bookendVars.startOfWorkDayMilli || adjustedDateMilli > bookendVars.endOfWorkDayMilli) { //since the bookendVars is in reference to the date variable, this function will still trigger if adjustedDate is technically within office hours, but on a different day

                //#region SETS ADJUSTMENT NUMBER VALUES ---------------------------------------------------------------------------------

                  var dayRemainder = (((bookendVars.endOfWorkDayMilli - dateMilli) / 1000) / 60) / 60; //time between end of work day and the original date time
                  var remainingAdjust = adjustmentNumber - dayRemainder; //gives us the remaining adjustment hours based off of what was already used to get to the end of the work day
                  var remainingAdjustMilli = remainingAdjust * 3600000;

                //#endregion ------------------------------------------------------------------------------------------------------------

                //#region NEW DAY CALCULATIONS ------------------------------------------------------------------------------------------

                  var newDay = new Date(addADay);

                  //adds remaining adjustment hours to the beginning of the work day the next day after date (addADay)
                  var dateTimeAdjusted = newDay.setMilliseconds((newDay.getMilliseconds() + remainingAdjustMilli));

                  var dateTimeAdjustedConvert = new Date(dateTimeAdjusted); //convert serial number to date object

                  date = dateTimeAdjustedConvert; //not sure if it should be date or something else yet. Need to make sure that the function works with this

                //#endregion ------------------------------------------------------------------------------------------------------------

                //#region SET LOOP VARIABLES IF STILL NOT WITHIN OFFICE HOURS OR EXCEEDS OFFICE HOURS OF NEXT DAY -----------------------

                    //if the new date exceeds the office hours of addADay, then do this...
                    if (dateTimeAdjusted > bookendAddedDate.endOfWorkDayMilli) {
                      adjustmentNumber = (remainingAdjust - addADayTitle.workDay) //subtracts remainingAdjust hours from the total workDay hours in the addADay variable
                      var dayAfterTomorrow = new Date(addADay);
                      var newDayAfterTomorrow = getNextDay(dayAfterTomorrow);
                      date = new Date(newDayAfterTomorrow.nextDay);
                      loop = true;
                      return {
                        date,
                        adjustmentNumber,
                        loop
                      };
                    } else {
                      loop = false;
                      return {
                        date,
                        //adjustmentNumber,
                        loop
                      };
                    };

                  //#endregion -------------------------------------------------------------------------------------------------------------
              
              } else {
                date = adjustedDate;
                loop = false;
                return {
                  date,
                  adjustmentNumber,
                  loop
                };
              };
            
            //#endregion --------------------------------------------------------------------------------------------------------------------

          };

        //#endregion ---------------------------------------------------------------------------------------------------------------------------


        //#region FIND WORK DAY FUNCTION -------------------------------------------------------------------------------------------------------

          /**
           * Returns the number of hours in a specific work day by subtracting the start from the end of the day, based on the properties loaded by the weekday variable
           * @param {Object} weekday A weekday variable with all it's associated properties
           * @returns Number
           */
          function findWorkDay(weekday) {

            //sets start time for weekday variable to a date for calculations
            var start = new Date(0); //69, baby
            start.setHours(weekday.startHour);
            start.setMinutes(weekday.startMinute);
            start.setSeconds(0);

            //sets end time for weekday variable to a date for calculations
            var end = new Date(0); //seriously though, just making sure the dates for both variables will always be the same
            end.setHours(weekday.endHour);
            end.setMinutes(weekday.endMinute);
            end.setSeconds(0);

            var workDayTime = (((end.valueOf() - start.valueOf()) / 1000) / 60) / 60; //subtracts end of day from start of day to get total work day hours for that weekday, then converts the milliseconds into hours (with decimal for minutes, if any)

            weekday.workDay = workDayTime; //sets our number to the variable 

            return weekday.workDay //returns our number to the actual object variable outside of the function

          };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------


        //#region DAY OF WEEK FUNCTION ---------------------------------------------------------------------------------------------------------

          /**
           * Returns a number 0-6 (Sunday - Saturday) based on the date input
           * @param {Date} d loads a date variable
           * @returns Number
           */
          function dayOfWeek(d) { //finds the day of the week
            var day = d.getDay();
            return day;
          };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------------


        //#region TITLE DAY OF WEEK FUNCTION ---------------------------------------------------------------------------------------------------

          /**
           * Returns the weekday variable, with all it's associated properties, from the weekday index input value
           * @param {Number} d The indexed number (0-6) of the weekday
           * @returns An object with properties
           */
          function titleDOW(d) { //returns the day of the week (refered to directly in another variable) based on the dayID index number
            if (d == 0) {
              return sunday;
            } else if (d == 1) {
              return monday;
            } else if (d == 2) {
              return tuesday;
            } else if (d == 3) {
              return wednesday;
            } else if (d == 4) {
              return thursday;
            } else if (d == 5) {
              return friday;
            } else if (d == 6) {
              return saturday;
            };
          };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------------


        //#region START/END/MIDNIGHT FUNCTION --------------------------------------------------------------------------------------------------
          
          /**
           * Sets start of work day, end of work day, and midnight to a date variable (including millisecond versions), returning an object with specific properties for each
           * @param {Date} originalDate A date variable (in this case, the date before any alterations)
           * @param {object} weekday A weekday variable with all its associated properties
           * @returns An object with properties
           */
          function startEndMidnight(originalDate, weekday) {

            var startOfWorkDay = new Date(originalDate); //adjusts start time of work day based on the day of the week
            startOfWorkDay.setHours(weekday.startHour);
            startOfWorkDay.setMinutes(weekday.startMinute);
            startOfWorkDay.setSeconds(0);
            var startOfWorkDayMilli = startOfWorkDay.getTime();

            var endOfWorkDay = new Date(originalDate); //adjusts end time of work day based on the day of the week
            endOfWorkDay.setHours(weekday.endHour);
            endOfWorkDay.setMinutes(weekday.endMinute);
            endOfWorkDay.setSeconds(0);
            var endOfWorkDayMilli = endOfWorkDay.getTime();

            var midnight = new Date(originalDate);
            midnight.setDate(midnight.getDate() + 1);
            midnight.setHours(0);
            midnight.setMinutes(0);
            midnight.setSeconds(0);
            var midnightMilli = midnight.getTime();

            return {
              startOfWorkDay,
              startOfWorkDayMilli,
              endOfWorkDay,
              endOfWorkDayMilli,
              midnight,
              midnightMilli
            };

          };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------------


        //#region GET NEXT DAY FUNCTION --------------------------------------------------------------------------------------------------------

          /**
           * Adds a day to the date variable and sets it to the start time of that new day's day of the week. Also adjusts for weekends if needed.
           * @param {Date} date A date object
           * @returns An object with properties
           */
          function getNextDay(date) {

            var nextDay = new Date(date);
            var newNextDay = nextDay.setDate(nextDay.getDate() + 1); //returns the day after the original date
            nextDay = new Date(newNextDay);
            var nextDayDayOfWeek = dayOfWeek(nextDay);
            var nextDayTitle = titleDOW(nextDayDayOfWeek); //returns a day title based on the dayID of the addADay variable

              if ((nextDayDayOfWeek == 6) || (nextDayDayOfWeek == 0)) { //checks if nextDay falls on a weekend
                nextDay = weekendAdjust(nextDay, nextDayDayOfWeek); //adjusts nextDay output to not fall on a weekend
                nextDayDayOfWeek = dayOfWeek(nextDay);
                nextDayTitle = titleDOW(nextDayDayOfWeek);
              };

              nextDay.setHours(nextDayTitle.startHour);
              nextDay.setMinutes(nextDayTitle.startMinute);
              nextDay.setSeconds(0);
              return {
                nextDay,
                nextDayTitle
              };
          };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------------


        //#region WEEKEND ADJUST FUNCTION ------------------------------------------------------------------------------------------------------
          
          /**
           * If input date falls on a weekend, returns a new date adjusted to start on the next upcoming Monday
           * @param {Date} date A date variable
           * @param {Number} dateWeekday A number indexed 0-6 representing the weekday of the date variable
           * @returns Date
           */
          function weekendAdjust(date, dateWeekday) {
            if (dateWeekday == 6) {
              var weekend = new Date(date);
              weekend.setDate(weekend.getDate() + 2);
              return weekend;
            } else if (dateWeekday == 0) {
              var weekend = new Date(date);
              weekend.setDate(weekend.getDate() + 1);
              return weekend;
            };
          };

        //#endregion ------------------------------------------------------------------------------------------------------------------------------


        //#region CONVERT DATE TO SERIAL ----------------------------------------------------------------------------------------------------------

          /**
           * Converts input date into serial number that excel can apply conditional formatting to
           * @param {Date} inDate A date variable
           * @returns String
           */
          function JSDateToExcelDate(inDate) {

            var returnDateTime = 25569.0 + ((inDate.getTime() - (inDate.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
            //var returnDateTime = 25569.0 + ((inDate.getTime()) / (1000 * 60 * 60 * 24));
            return returnDateTime.toString().substr(0,20);
        
          };

        //#endregion --------------------------------------------------------------------------------------------------------------------------------

      //#endregion -------------------------------------------------------------------------------------------------------------------------------

  //#endregion -------------------------------------------------------------------------------------------------------------------------------------


//#endregion ---------------------------------------------------------------------------------------------------------------------------------------



//#region ERROR HANDLING ------------------------------------------------------------------------------------------

  //#region TRY CATCH ---------------------------------------------------------------------------------------------
  async function tryCatch(callback) {
    try {
      await callback();
    } catch (error) {
      console.error(error);
    }
  }
//#endregion ---------------------------------------------------------------------------------------------------

//#endregion -----------------------------------------------------------------------------------------------------