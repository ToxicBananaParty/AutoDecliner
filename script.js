
// ----------- USER INPUTS ---------------- //

var myEmail = "name@example.com";

// Days of the week you want to decline events for
// 0 for Sunday, 1 for Monday, 2 for Tuesday, and so on.
// For Most People: 0 for Sundays and 6 for Saturdays
var declineDays = [3, 6];

// Your working hours
// These must be put in manually because the Calendar API does not have any way to access Working Hours information.
// As they continue dev on their "Calendar User Availability API" we may be able to update this to automatically get working hours.
var startTime = 9; // 9:00am
var endTime = 17.5; // 5:30pm

var autoResponse = CalendarApp.GuestStatus.MAYBE; // Choose between [MAYBE, NO]

// ----------- END USER INPUTS -------------- //


function dayToString(day) {
    switch(day) {
    case 0:
        return "Sunday";
    case 1:
        return "Monday";
    case 2:
        return "Tuesday";
    case 3:
        return "Wednesday";
    case 4:
        return "Thursday";
    case 5:
        return "Friday";
    case 6:
        return "Saturday";
    }
    
    return "";
}

function convertNumToTime(number) {
    number = Math.abs(number);
    
    // Separate the hour from the minutes
    var hour = Math.floor(number);
    var remainder = (1/60) * Math.round((number - hour) / (1/60));
    var minute = Math.floor(remainder * 60) + "";
    
    // Add leading '0' if single-digit minute
    if (minute.length < 2) {
        minute = "0" + minute; 
    }
    
    // Convert to 12-hour time
    if(hour > 12) {
        return (hour - 12) + ":" + minute + "pm";
    } else {
        return hour + ":" + minute + "am";
    }
    
}


function autoDeclineOutsideWorkingHours() {
    var calendar = CalendarApp.getCalendarById(myEmail);
    
    var today = new Date();
    var twoWeeks = new Date();
    twoWeeks.setDate(today.getDate() + 14);
    
    var events = calendar.getEvents(today, twoWeeks);
    for (var i = 0; i < events.length; i++) {
        var eventStart = events[i].getStartTime().getHours() + (events[i].getStartTime().getMinutes() / 60);
        var eventEnd = events[i].getEndTime().getHours() + (events[i].getEndTime().getMinutes() / 60);
        var eventStatus = events[i].getMyStatus(); // Get the status of the event for the current user
        var eventDay = events[i].getStartTime().getDay(); // get the day of the week for the event
        var eventTitle = events[i].getTitle();
        
        // Check if the event is outside working hours
        if (eventStart < startTime || eventEnd > endTime) {
            if(eventStatus == CalendarApp.GuestStatus.INVITED) {
                Logger.log(`Declining event "${eventTitle}" on ${events[i].getStartTime().toDateString()} from ${convertNumToTime(eventStart)} - ${convertNumToTime(eventEnd)} because it is outside working hours`);
                events[i].setMyStatus(autoResponse);
            } else {
                Logger.log(`Would decline event "${eventTitle}" on ${events[i].getStartTime().toDateString()} from ${convertNumToTime(eventStart)} - ${convertNumToTime(eventEnd)} because it is outside working hours, but it has already been responded to.`);
            }
            
            continue; // Don't bother checking non-working day since we've already responded to this event now
        }
        
        // Check if the event is on a non-working day
        if (declineDays.indexOf(eventDay) != -1) {
            var day = declineDays[declineDays.indexOf(eventDay)];
            if(eventStatus == CalendarApp.GuestStatus.INVITED) {
                Logger.log(`Declining event "${eventTitle}" on ${events[i].getStartTime().toDateString()} from ${convertNumToTime(eventStart)} - ${convertNumToTime(eventEnd)} because it is on a non-working day ${dayToString(day)}`);
                events[i].setMyStatus(autoResponse);
            } else {
                Logger.log(`Would decline event "${eventTitle}" on ${events[i].getStartTime().toDateString()} from ${convertNumToTime(eventStart)} - ${convertNumToTime(eventEnd)} because it is on a non-working day ${dayToString(day)}, but it has already been responded to.`);
            }
        }
    }
}

function attachTrigger() {
    ScriptApp.newTrigger("autoDeclineOutsideWorkingHours").forUserCalendar(myEmail).onEventUpdated().create();
}
