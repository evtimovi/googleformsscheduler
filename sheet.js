/**
  * Generates the menu that will have an option to call the update script.
    */
function onOpen(e){
    Logger.clear();
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Schedule Updater")
        .addItem('Add events from calendar', 'addFromCalendar_')
        .addItem('Populate form', 'populateForm_')
        .addItem('Clear form', 'clearForm_')
        .addToUi();

}

function clearForm_()
{
    var form = FormApp.openByUrl('https://docs.google.com/a/lafayette.edu/forms/d/1qG2-gfXGMgcYsJBWWFv5YkDzsA4PG6Z8Jtzo1dwaC2g/edit?usp=drive_web');
    form.getItems().forEach(function(e){ form.deleteItem(e)});
    //var pageIndices = pageBreaks.map(function(e){return e.getIndex()});
    //pageIndices.forEach(function(e){form.deleteItem(e)});
}

function populateForm_()
{

    var calendarSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Signups"); //get the calendar sheet

    //get the headers - this returns a JS array of Date objects signifying Midnight on the day of
    var days = calendarSheet.getRange(1, 2, 1, calendarSheet.getLastColumn() - 1).getValues()[0];

    var times = calendarSheet.getRange(2, 1, calendarSheet.getLastRow() - 1, 1)
        .getValues().map(function(e){return new Date(e); });

    var scheduleGrid = calendarSheet.getRange(2, 2, calendarSheet.getLastRow() - 1, calendarSheet.getLastColumn() - 1);

    var form = FormApp.openByUrl('https://docs.google.com/a/lafayette.edu/forms/d/1qG2-gfXGMgcYsJBWWFv5YkDzsA4PG6Z8Jtzo1dwaC2g/edit?usp=drive_web');

    var pages = [];

    days.forEach(function(day, index){
        var dayString = day.toDateString().slice(0,10);

        var page = form.addPageBreakItem().setTitle(dayString);
        pages[day] = (page);

        var question = form.addMultipleChoiceItem().setTitle("When are you leaving on " + dayString + '?');
        question.setRequired(true);
        question.setHelpText("Pick a time to check out.\nIf you got here by mistake and want to select another day, click the 'Back' button.\nIf no of these times work for you, please contact me ASAP.");

        //get the column corresponding to the date
        var currColumn = scheduleGrid.offset(0, index, scheduleGrid.getNumRows(), 1);

        var choices = [];

        for(var timeIndex = 1; timeIndex < currColumn.getNumRows(); timeIndex++)
        {
            if(currColumn.getCell(timeIndex, 1).getValue() != 'Unavailable')
            {
                choices.push(question.createChoice(times[timeIndex-1].toTimeString().slice(0,5)));
            }
        }
        question.setChoices(choices);
        page.setGoToPage(FormApp.PageNavigationType.SUBMIT);
    });

    var formsPage = form.addPageBreakItem().setTitle("General Information");

    var nameField = form.addTextItem().setTitle("What is your name?").setRequired(true);
    var roomField = form.addTextItem().setTitle("What is your room number in Gates hall?").setRequired(true);

    var mc = form.addMultipleChoiceItem();
    mc.setChoices(days.map(function(e){ return mc.createChoice(e.toDateString().slice(0,10), pages[e]); }));
    mc.setTitle("Which day are you leaving on?");
    mc.setRequired(true);

    form.addPageBreakItem().setTitle("unreachable");
    // grid.setRows(times.map(function(e){ return e.toTimeString().slice(0,5); }));
}

function addFromCalendar_()
{
    var calendarSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Signups"); //get the calendar sheet

    //get the headers - this returns a JS array of Date objects signifying Midnight on the day of
    var days = calendarSheet.getRange(1, 2, 1, calendarSheet.getLastColumn() - 1).getValues()[0];

    var times = calendarSheet.getRange(2, 1, calendarSheet.getLastRow() - 1, 1)
        .getValues().map(function(e){return new Date(e); });

    var scheduleGrid = calendarSheet.getRange(2, 2, calendarSheet.getLastRow() - 1, calendarSheet.getLastColumn() - 1);
    var calendar = CalendarApp.getCalendarsByName("Ivan Evtimov")[0];

    //clear the grid
    scheduleGrid.clear({contentsOnly: true});

    //for each column:
    // 1. look up its day
    // 2. get the events between 9 am and 10:15 pm on this day
    // 3. select:
    //    - the first cell that this event hits
    //    - the last cell that this event hits
    // 4. merge the selected range, set color to blue, and type busy

    days.forEach(function(day, currDayColumn){
        //steps 1 and 2: get the events between 9 am and 10:15 pm on the day
        events = calendar.getEvents(new Date(day.setHours(times[0].getHours())), new Date(day.setHours(22,15)));

        //find the first cell for which the time it contains is greater than the current event's start time
        events.forEach(function(currEvent){
            var startTimeRow = -1;

            for(var i = 0; i < times.length; i++)
            {
                currTime = times[i];

                //event starts before first time on the grid
                if (currEvent.getStartTime().getHours() < times[0].getHours())
                {
                    startTimeRow = 2;
                    break;
                }

                //regular case
                if (currTime.getHours() >= currEvent.getStartTime().getHours() 
                        && currTime.getMinutes() >= currEvent.getStartTime().getMinutes())
                {
                    startTimeRow = i+2;
                    break;
                }
            }; 

            //not adding +1 because the index found will be the one of the next cell
            var endTimeRow = -1;
            for(var i = 0; i < times.length; i++)
            {
                currTime = times[i]; 

                //event ends after lsat time on the grid
                if (currEvent.getEndTime().getHours() >= times[times.length - 1].getHours()
                        || currEvent.getEndTime().getDate() > day.getDate())
                {
                    endTimeRow = calendarSheet.getLastRow()+1;
                    break;
                }
                if(currTime.getHours() >= currEvent.getEndTime().getHours() 
                        && currTime.getMinutes() >= currEvent.getEndTime().getMinutes())
                {
                    endTimeRow = i + 2;
                    break;
                }
            }

            //note that currDayColumn is incremented by 2 because the first column is the times
            var eventCellsRange = calendarSheet.getRange(startTimeRow, currDayColumn + 2, endTimeRow - startTimeRow);
            eventCellsRange
                //.getCell(1,1)
                .setValue("Unavailable")
                .setBackgroundRGB(36,89,237)
                .setVerticalAlignment("middle")
                .setFontSize(11)
                .setFontColor("white")
                .setFontWeight("bold")
                .setHorizontalAlignment("center")
                .setWrap(true);
            // eventCellsRange.merge();
        });
    });
}

