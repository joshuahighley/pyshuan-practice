const auditorInfo = [
  ['#c9daf8', testEmail+'+6@institution.org'],
  ['#d0e0e3', testEmail+'+7@institution.org'],
  ['#d9ead3', testEmail+'+8@institution.org']
];



// Adds the script to a new dropdown menu in Sheets:
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Report Generation')
    .addItem('Send Auditor To-Dos','getRowsByColumnColor')
    .addToUi();
}


// The thing that sends the emails:
function getRowsByColumnColor() {
  // Get the active spreadsheet and sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = spreadsheet.getSheetByName('Auditing');
  
  // Get the data range and values
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  // Get the background colors of column C
  var colors = sheet.getRange(1, 3, sheet.getLastRow(), 1).getBackgrounds();
  
  // Define the target colors
  var targetColors = [
    "#c9daf8",  // light cornflower blue 3
    "#d0e0e3",  // light cyan 3
    "#d9ead3"   // light green 3
  ];


  // Filter and group rows based on column C color
  var rowsByColor = {};
  
  values.forEach((row, index) => {
    var cellColor = colors[index][0];
    for (var colorName in targetColors) {
      if (targetColors[colorName] === cellColor) {
        if (!rowsByColor[colorName]) {
          rowsByColor[colorName] = [];
        }
        rowsByColor[colorName].push(row);
        break;
      }
    }
  });
  

  const uniqueClients = {};

  for (let group in rowsByColor) {
    // Sort each group by the first element in each row (the "todo" ID number)
    rowsByColor[group].sort((a, b) => a[0] - b[0]);

    // Loop through each row in the group and add the ID (first of each row) to a new set
    // then convert the set to an array and store it in uniqueClients object
    const uniqueIds = new Set();

    rowsByColor[group].forEach(row => {
        uniqueIds.add(row[0]);
    });

    uniqueClients[group] = Array.from(uniqueIds);
  }

  console.log('unique IDs')
  console.log(uniqueClients);


  // From Perplexity
  function concatenateTodos(uniqueClients, rowsByColor) {
    let todosByGroup = {}; // This will hold concatenated todos by group

    for (let group in rowsByColor) {
        let todosSort = {};

        // Loop through each row in the current group to organize todos by client ID
        rowsByColor[group].forEach(row => {
            let clientNumber = row[0];
            let name = row[1];
            let description = row[2];
            let status = row[19];  // Assuming status is at index 20

            // Initialize entry for client if it doesn't exist
            if (!todosSort[clientNumber]) {
                todosSort[clientNumber] = { name: name, descriptions: [], status: status };
            }

            // Add the description to the list of descriptions for this client
            todosSort[clientNumber].descriptions.push(description);
        });

        // Create the final array structure for each group
        todosByGroup[group] = uniqueClients[group].map(clientNumber => {
            let clientData = todosSort[clientNumber] || { name: '', descriptions: [], status: '' };
            let concatenatedDescription = clientData.descriptions.join('<br>');
            return [clientNumber, clientData.name, concatenatedDescription, clientData.status];
        });
    }

    return todosByGroup;
  }


  let todosByClient = concatenateTodos(uniqueClients, rowsByColor);
  
  console.log('To-Dos By Client:')
  console.log(todosByClient); 


  // Create email subject, report date, header, then report body with auditors' to-dos
  // Date example: "Friday, Jul 2, 2021"  year:"numeric",
  reportDate = new Date().toLocaleDateString('en-us', { weekday:"long", month:"short", day:"numeric"});
  reportSubject = "Open To-Dos for " + reportDate
  reportHeader = '<p>Auditor, these are your open to do\'s as of ' + reportDate + ':</p>';
  var todoBody = ['','',''];

  for (let group in todosByClient) {
      todosByClient[group].forEach(row => {
        let hhLine = 'HH: ' + row[0] + ', ' + row[1]
        let todoLine = row[2]
        let pathLine = 'Path: <i>' + row[3] + '</i>'

        todoBody[group] += '<p>' + hhLine +'<br>'+ todoLine +'<br>'+ pathLine + '</p>'

    });
  }


  // Send email with To Do report
  for (let auditor in auditorInfo) {
    email = auditorInfo[auditor][1]
    GmailApp.sendEmail(email, reportSubject, '', {htmlBody: reportHeader + todoBody[auditor]})
    };

}
