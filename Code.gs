/**
 * Adding events on Calendar using Forms
 */
function addTaskEvents() {

  var myCal = CalendarApp.getCalendarById('k3p6jh50sh6pcv05sgkhi26t8s@group.calendar.google.com'); // get calendar ID
  var mySheet = SpreadsheetApp.getActiveSheet(); // get sheet 
  var data = mySheet.getDataRange().getValues(); // get sheet data

  for(var i = 1; i < data.length; i++){
     if(data[i][6] == ""){ // if isDone is empty, execute below

       /* Set date and time(Start) */
      var eventStartDateTime = new Date(data[i][3]);
      eventStartDateTime.setHours(eventStartDateTime.getHours());
      eventStartDateTime.setMinutes(eventStartDateTime.getMinutes());   

       /* Set duration of lesson */
      var duration = data[i][4].replace(/[^0-9]/g, ''); // レッスンの時間を取得(数字)
      var durationInt = parseInt(duration, 10); // str -> int

      /* Set time(End) */
      var endDate = new Date(data[i][3]);
      endDate.setMinutes(eventStartDateTime.getMinutes() + durationInt); // start + duration(30 or 60)
      var eventEndDateTime = new Date(endDate); // set

      /* Set title of Calendar */
      var eventTitle = '[' + data[i][1] + ' 様' + '] Yuki\'s English Lesson';
      // console.log(eventTitle);
      // console.log('start ' + eventStartDateTime);
      // console.log('end ' +eventEndDateTime);

      /* Set other options
          guests: email address from Form
          sendInvites：whether sending invitaion to the email address from Form (T/F)
      */

      var showMinute = "00";
      if (eventStartDateTime.getMinutes() != 0) {
        showMinute = eventStartDateTime.getMinutes();
      }


      var option = {
        guests:data[i][2],  // アドレス
        sendInvites: true,
        description: 
        `---------------------------------\n\n` + 
        data[i][1] + `, thank you for your reservation!‪✩‬‪✩‬‪✩‬ \n` +
        `\n [Your Google Form Details] ` + 
        `\n Name: ` + data[i][1] +
        `\n Email address : ` + data[i][2] +
        `\n Lesson date (JST) : ` + eventStartDateTime.getFullYear() + `/` + (eventStartDateTime.getMonth() + 1) + `/` + 
          eventStartDateTime.getDate() + ` ` + eventStartDateTime.getHours() + `:` + 
          showMinute + 
        `\n Length of lesson : ` + data[i][4] + 
        `\n Any questions : ` + data[i][5] + 
        `\n\n Looking forward to talking to you! \n\n Sincerely, \n Yuki` +
        `\n\n---------------------------------` 
      }

      /* add this event to Calendar */
      myCal.createEvent(eventTitle,eventStartDateTime,eventEndDateTime, option); // createEvent(title, start, end, option)
      /* update isDone cel TRUE on spreadsheet */
      mySheet.getRange(i + 1,7).setValue(true); 
    }
  }
}

////////////////////////////////////////////////////////////////////////
// memo 1: How to show appscript.json
// 1. Click the gear icon(⚙) on the left 
// 2. Check "Show "appsscript.json" manifest file in editor"

// memo 2: Language of invitation email
// It depends on which language the email receiver uses on Google Calendar
