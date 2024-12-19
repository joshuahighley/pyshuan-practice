function sendTodoReports() {

  // Get the active spreadsheet and sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('To Do');
  var lastRow = ss.getLastRow();

  // Removes empty cells from lists, combines with ', ' for better legibility in emails
  function listCleaner(rawList) {
    return rawList.map(subArray => subArray.filter(item => item !== '').join(', '));
    };

  // Collects all cells in the  To Do column, blanks included
  var todos = sheet.getRange(2, 3, lastRow-1, 1).getValues();
  var todoHHNs = listCleaner(sheet.getRange(2, 1, lastRow-1, 2).getValues());
  var todoTeams = listCleaner(sheet.getRange(2, 22, lastRow-1, 4).getValues());

  // Sorts a given list based on if the todo at the same index contains the given name
  function sortList(name, list) {
    var result = []
    for(let i = todos.length-1; i >= 0; i--) {
      if(todos[i].toString().includes(name)) {
        result.push(list[i]);
      }
    }
    return result
  }

  // Creates an email with the given email address, team member's name, and todo list
  function emailWriter(emailAddress, personName, todos) {
    var date = new Date().toLocaleString('en-us', { weekday:"long", month:"short", day:"numeric"});
    var time = new Date().toLocaleString('en-us', { hour:"numeric", minute:"numeric"});
    var subject = "Open To-Dos for " + date;
    let header = '<p>' + personName + ',<br><br>These are your open to-do\'s on the Nexus as of today at ' + time + ':</p>';
    let noTodos = personName + ',<br><br>You have no open to-do\'s on Nexus as of ' + time + '.';
    let warning = '<br><br><br><i>This is the first test run of the report generator script. Please let me know if you find discrepancies between this report and the Nexus. Thanks for your patience.</i>'

    if (todos == '') {
      GmailApp.sendEmail(emailAddress, subject, '', {htmlBody: noTodos + warning});
    } else {
      GmailApp.sendEmail(emailAddress, subject, '', {htmlBody: header + todos + warning});
    };
  };

  // Checks each todo for a person's name, groups into their email body, sends emails with emailWriter function
  teamInfo.forEach((person, name) => {
    var todosCombo = []
    var a = sortList(name, todoHHNs)
    var b = sortList(name, todoTeams)
    var c = sortList(name, todos)

    for(let i = c.length-1; i >= 0; i--){
      let todoBoo = '<p><br><b>HH:' + a[i] +'</b><br>'+ b[i] +'<br>'+ c[i] + '</p>'
      todosCombo.push(todoBoo)
    };
    
    emailWriter(person.email, name, todosCombo.join(''))
  });

}
