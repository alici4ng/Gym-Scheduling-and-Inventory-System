
// This code is the main code that updates when current quantity of stocks fall below a certain value


function onEdit(e) {
    // Color Hex code
    let color = "#ff0000"; // Red
    
    try {
      if (!e) {
        e = {
          source: SpreadsheetApp.getActiveSpreadsheet(),
          range: SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell()
        };
      }
  
      var ui = SpreadsheetApp.getUi();
      var editedRange = e.range;
      var sheet = editedRange.getSheet();
      var row = editedRange.getRow();
      var col = editedRange.getColumn();
      var cellKey = sheet.getName() + "!" + row + ":" + col;
      
      // Get edited value
      var editedValue = editedRange.getValue();
  
      var scriptProperties = PropertiesService.getScriptProperties();
      var previousValue = scriptProperties.getProperty(cellKey);
      
      if (previousValue === null) {
        previousValue = editedValue; 
      } else {
        // Compare previous value with current value
        if (isColorChanged(editedRange, color)) {
          var editTime = new Date();
          updateCalendar(previousValue, editedValue, editTime);
        }
      }
      
      // Store the current value as previous 
      scriptProperties.setProperty(cellKey, editedValue);
  
      // Check if the edited cell is in the 'Current Quantity' column (e.g., column B)
      var currentQuantityColumn = 2;
      
      if (col == currentQuantityColumn) {
        var currentQuantity = editedValue;
        var reorderPoint = editedRange.offset(0, 1).getValue();  // Assuming reorder point is in the next column
        
        if (currentQuantity <= reorderPoint) {
          var itemName = editedRange.offset(0, -1).getValue();  // Assuming item name is in the previous column
          // Show only the necessary alert
          ui.alert('Restock Alert', 'The item "' + itemName + '" needs to be restocked. Current Quantity: ' + currentQuantity , ui.ButtonSet.OK);
        }
      }
  
      // Notify Edit by user email 
      
      var cell = editedRange.getA1Notation();
      var newValue = editedValue; // Use editedValue instead of e.value
      var email = Session.getActiveUser().getEmail();
  
      // Store the notification message in PropertiesService for later retrieval
      PropertiesService.getDocumentProperties().setProperty('notificationMessage', 
        'Cell ' + cell + ' in sheet "' + sheet.getName() + '" was changed to "' + newValue + '" by ' + email + '.');
  
      // Show the notification bubble immediately
      // Ensure this alert only appears if necessary
      if (col != currentQuantityColumn) {
        ui.alert('Cell ' + cell + ' in sheet "' + sheet.getName() + '" was changed to "' + newValue + '" by ' + email + '.');
      }
  
    } catch (error) {
      Logger.log("Error in onEdit:", error);
    }
  }
  
  
  function updateCalendar(iniValue, ediValue, time) {
    try {
      // Calendar ID (Look in ur calendar settings and change to that)
      var calendarId = '53bda047c5d81c36b3e2037ef4b352e0838022f4e5c2bfbcacf9345abcf2bd96@group.calendar.google.com'; 
      var eventTitle = 'Sheet Edit Event';
      var eventDescription = 'Value changed from: ' + iniValue + ' to ' + ediValue;
      
      var event = {
        summary: eventTitle,
        description: eventDescription,
        start: {
          dateTime: Utilities.formatDate(time, "UTC+8", "yyyy-MM-dd'T'HH:mm:ss'Z'")
        },
        end: {
          dateTime: Utilities.formatDate(time, "UTC+8", "yyyy-MM-dd'T'HH:mm:ss'Z'")
        }
      };
      
      // Create event in Calendar
      var createdEvent = Calendar.Events.insert(event, calendarId);
      
      Logger.log('Event created: %s', createdEvent.htmlLink);
    } catch (error) {
      console.error("Error in updateCalendar:", error);
    }
  }
  
  function isColorChanged(range, color){
    // Get cell color
    var editedBackgroundColor = range.getBackground();
      if (editedBackgroundColor === color) {
        return true;
      }
      else{
        return false;
      }
  }
  
  function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Notification')
      .addItem('Show Notification', 'showNotification')
      .addToUi();
  }
  
  function showNotification() {
    var ui = SpreadsheetApp.getUi();
    var message = PropertiesService.getDocumentProperties().getProperty('notificationMessage');
    ui.alert(message || 'No recent notifications.');
  }
  
  function testUpdateCalendar() {
    var value = "Test Value";
    var time = new Date(); 
  
    // Call updateCalendar function
    updateCalendar(value, time);
  }
  
  