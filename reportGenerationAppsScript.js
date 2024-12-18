// Adds the script to a new dropdown menu in Sheets:
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Report Generation')
    .addItem('Send To-Do Reports','sendTodoReports')
    .addToUi();
}

// Staff email Map ommitted

function sendTodoReports() {

  // Get the active spreadsheet and sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('To Do');
  var lastRow = ss.getLastRow();

  // Collects all cells in the  To Do column, blanks included
  var todos = sheet.getRange(2, 3, lastRow-1, 1).getValues();
  var todoHHNs = sheet.getRange(2, 1, lastRow-1, 2).getValues();
  var todoTeams = sheet.getRange(2, 22, lastRow-1, 4).getValues();


  var todosCombined = []

  for(let i = todos.length-1; i >=0; i--) {
    
    var todoCombination = ['']
    
    if(todos[i] != '') {
      let a = todoHHNs[i][0] + ', ' + todoHHNs[i][1]
      let b = todoTeams[i]
      let c = todos[i]
      todoCombination += '<p><b>HH:' + a +'</b><br>'+ b +'<br>'+ c + '</p>'
    };
    todosCombined.push(todoCombination)
  };

  

  // Creates an email with the given email address, team member's name, and todo list
  function emailWriter(emailAddress, personName, todos) {
    var date = new Date().toLocaleString('en-us', { weekday:"long", month:"short", day:"numeric"});
    var time = new Date().toLocaleString('en-us', { hour:"numeric", minute:"numeric"});
    var subject = "Open To-Dos for " + date;
    let header = '<p>' + personName + ', these are your open to-do\'s as of today at ' + time + ':</p>';
    let body = {htmlBody: header + todos};

    GmailApp.sendEmail(emailAddress, subject, '', body);
  };

  
  // Checks each todo for a person's name, groups into their email body, sends emails with emailWriter function
  teamInfo.forEach((person, name) => {

    var todosByName = []

    todosCombined.forEach(todo => {
      if (todo.includes(name)) {
        todosByName.push(todo)
      }
    });

    // Uses the function from above to create and send emails with the to-do groups and team info
    emailWriter(person.email, name, todosByName.join(''))
    
  });

} //end of line - FRESH START
