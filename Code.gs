// Function to create a custom menu on opening the spreadsheet
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Drive Access')
        .addItem('Get File Name', 'GetFileNameByUrl')
        .addToUi();
  }
  
  // Main function to get the file names
  function GetFileNameByUrl(){
    // Prompt user to enter the link range
    ranges = Browser.inputBox("Enter link range", Browser.Buttons.OK_CANCEL);
  
    // Access the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  
    // Trim the arguments and split the input range and output range
    var args = ranges.split(',').map(arg => arg.trim());
    var inputRange = args[0];
    var outputRange = args[1];
  
    // Get the active sheet and updated input range
    var sheet = ss.getActiveSheet();
    var inputRange = sheet.getRange(inputRange);
    
    // Get the number of rows in the input range
    var numRows = inputRange.getNumRows();
  
    // Get the values from the input range
    var rangeValues = inputRange.getValues();
  
    // Initialize an array to store the document names
    var docName = [];
    for (var row = 0; row < numRows; row++) {
      // Extract the file ID from the row values
      var fileId = String(rangeValues[row]).match(/[-\w]{25,}/i);
      
      // If a valid file ID is found, fetch the document and store its name
      if (fileId) {
        var document = DriveApp.getFileById(fileId[0]);
        docName.push(document.getName());
      }
      // If no file is found, add "No Document Found" to the array
      else {
        docName.push(["No Document Found"]);
      }
    }
  
    // Convert the array to a 2D array for output
    var outputValues = ConvertTo2DArray(docName);
  
    // Set the values to the specified output range
    sheet.getRange(outputRange).setValues(outputValues);
  }
  
  // Function to convert a 1D array to a 2D array
  function ConvertTo2DArray(arr) {
    var newArr = [];
    for (var i = 0; i < arr.length; i++) {
      newArr.push([arr[i]]);
    }
    return newArr;
  }
  