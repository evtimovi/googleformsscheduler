function onFormSubmit(e) {


    var response = e.response;

    Logger.log("response: " + response);
    Logger.log("item responses: " );
    response.getItemResponses().forEach(function(e, i){
        Logger.log("response " + i + " title: " + e.getItem().getTitle()); 
    });

    var name = e.response.getItemResponses()[1].getResponse();
    name += " (" + e.response.getItemResponses()[2].getResponse() +")";

    var dateAndTime = e.response.getItemResponses()[3].getResponse();
    var time = e.response.getItemResponses()[0].getResponse();
    dateAndTime += " " + time;

    Logger.log("respondent's name and room number: " + name);
    Logger.log("respondent's date and time: " + dateAndTime);

    handleCalendarEntry(name, dateAndTime);

    var timeQuestionItem = e.response.getItemResponses()[0].getItem().asMultipleChoiceItem();

    timeQuestionItem.setChoices(timeQuestionItem.getChoices().filter(function(e){ return e.getValue() !== time; }));

}

function handleCalendarEntry(name, dateAndTime)
{
    var calendarSheet = SpreadsheetApp.openByUrl("https://docs.google.com/a/lafayette.edu/spreadsheets/d/1txk2WqRk5l01mG1OGITLhofyLxwrPhTvyEOY0ztZgHg/edit?usp=drive_web").getSheetByName("Signups");
    var scheduleGrid = calendarSheet.getRange(2, 2, calendarSheet.getLastRow() - 1, calendarSheet.getLastColumn() - 1);
    var days = calendarSheet.getRange(1, 2, 1, calendarSheet.getLastColumn() - 1).getValues()[0];
    var times = calendarSheet.getRange(2, 1, calendarSheet.getLastRow() - 1, 1)
        .getValues().map(function(e){return new Date(e); });

    //Logger.log("dateAndTime: " + Date.parse(dateAndTime));
    var apptDate = new Date();
    apptDate.setHours(parseInt(dateAndTime.slice(11,13)));
    apptDate.setMinutes(parseInt(dateAndTime.slice(14,17)));
    apptDate.setDate(parseInt(dateAndTime.slice(9,11)));
    Logger.log("date and time: " + apptDate.toString());

    var dayColumn = -1;
    var timeRow = -1;

    for(var i = 0; i < times.length; i++)
    {
        currTime = times[i];

        //regular case
        if (currTime.getHours() === apptDate.getHours() 
                && currTime.getMinutes() === apptDate.getMinutes())
        {
            timeRow = i + 1;
            break;
        }
    }; 

    for(var i = 0; i < days.length; i++)
    {
        currDay = days[i];

        //regular case
        if (currDay.getDate() === apptDate.getDate())
        {
            dayColumn = i + 1;
            break;
        }
    }; 

    scheduleGrid.getCell(timeRow, dayColumn).setValue(name)
        .setBackgroundRGB(237,171,28)
        .setVerticalAlignment("middle")
        .setFontSize(11)
        .setFontColor("black")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setWrap(true);

    var calendar = CalendarApp.getDefaultCalendar();
    calendar.createEvent('Checking out ' + name,
            apptDate,
            new Date(apptDate.setMinutes(apptDate.getMinutes()+15)),
            {
                guests: 'ivanevtimov5@gmail.com',
                sendInvites: true
            });

}

