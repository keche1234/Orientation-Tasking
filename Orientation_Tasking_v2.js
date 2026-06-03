/*******************************************************************************
 * Original Script: Ellia Yang (SCS '25 / Orientation 2023)
 * Contributors:  Kenechukwu Echezona (SCS '26 / Orientation 2025)
 ******************************************************************************/
// When you open the spreadsheet, "onOpen" modifies the UI to have a button called Orientation
// Instructions on how to run in the comment for the onOpen

var calendarId = 'cmuorientationtasking@andrew.cmu.edu';

/*
 * These values can be found in the "Roster" tab.
 * Keep each position contiguous between [roleMin, roleMax)
 * THESE ARE 1-INDEXED (aka the front-end values)
 * When using these for array indices in code, subtract 1
 */
var hocMin = 4; // change the <role>Min/<role>Max values ONLY
var hocMax = 12;
var olMin = 12;
var olMax = 37;
var ocMin = 37;
var ocMax = 158;

var ocStart = 1;
var olStart = 1;
var hocStart = 1;

/*
 * These values refer to the columns in TrainingTasks,
 * BigTasks, RandomTasks, and ManualTasks
 */
var eventSubjectCol = 0;
var eventStartCol = 3;
var eventEndCol = 6;
var eventAttachmentCol = 7;
var eventDescriptionCol = 8;
var eventLocationCol = 9;
var eventIDCol = 10;
var eventFinalTitleCol = 11;
var ocsNeededCol = 12;
var olsNeededCol = 13;
var hocsNeededCol = 14;

/*
 * These values refer to the rows in the Tasks sheets
 */
var firstTaskRow = 0;

/*
 * Refer to the columns in the roster and assignment sheets
 * This is 0-INDEXED
 */
var rosterEmailCol = 1;

/*
 * These values are for whatever sheet will have the random schedules placed
 */
var assignmentEmailCol = 1; // offset of 0
var assignmentStartCol = 2;

/*
 * These values are for defining a range when writing to a schedule
 */
var assignmentStartLetter = "C";
var assignmentStartNum = 2; // offset of 1

/*
 * This value is for whatever column in the
 * "Roster" tab has a staff member's position
 */
var rosterPosLetter = "F";

var scheduleBuffer = 15 * 60; // staff should have a break of at least this time (in seconds) in between events

/*********************************************************
 * Runs... well, when the spreadsheet opens
 * Adds the buttons to the spreadsheet UI
 * 
 * Quic Guide
 * 1) Run the "GCal Creation/Setup" functions first.
 * 2) Add the manual tasks to the spreadsheet that will
 *    have random assignments written to, so the schedule
 *    generator will take into account those conflicts.
 * 3) Generate Random Schedules.
 * 4) Use the "Invite" functions (you can do the big invites earlier)
 ********************************************************/
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Orientation')
      // GCal Creation/setup
      .addItem('Create Selected Training Tasks', 'createTrainingEvents')
      .addItem('Invite All to Selected Training Events', 'sendTrainingInvites')
      .addSeparator() // Create based on selection
      .addItem('Create Selected Big GCal Events', 'createBigEvents')
      .addItem('Create Selected Manual GCal Events', 'createManualEvents')
      .addItem('Create Selected Random GCal Events', 'createRandomEvents')
      .addSeparator()
      // Random events to add to specific
      .addItem('Assign Random Tasks', 'assignRandomSchedules')
      .addSeparator()
      // Invitation process
      .addItem('Invite All Staff to Selected Big GCal Events', 'sendBigInvites')
      .addItem('Invite Selected OCs to Specific GCal Events (Manual + Random)', 'sendSelectedOCsInvites')
      .addItem('Invite Selected OLs to Specific GCal Events (Manual + Random)', 'sendSelectedOLsInvites')
      .addItem('Invite Selected HOCs to Specific GCal Events (Manual + Random)', 'sendSelectedHOCsInvites')
      //.addSeparator()
      //TODO: addItem('(DANGER) Delete Selected Events (GCcal)', 'delete_events()')
      .addToUi();
}

/*****************************************************************************
 *****************************************************************************
 * A: CREATING GCAL EVENTS
 * These functions use a provided schedule (of either "random," "manual," or
 * "big" tasks, the former two are "specific") to create GCal events, and
 * providing the shet with GCal Event IDs. 
 *****************************************************************************
 *****************************************************************************/

/*********************************
 * Creates Google Calendar Events
 *********************************/
function createTrainingEvents() {
  createTasksFromSheet_('TrainingSchedules');
}

function createBigEvents() {
  createTasksFromSheet_('BigTasks');
}

function createRandomEvents() {
  createTasksFromSheet_('RandomTasks');
}

function createManualEvents() {
  createTasksFromSheet_('ManualTasks');
}

function createTasksFromSheet_(sheetName) {
  let ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(sheetName);
  let range = sheet.getDataRange();
  let values = range.getValues();

  let activeRange = sheet.getActiveRange();
  if (activeRange) {
    setUpCalendarTasks_(values, range, activeRange);
    let startRow = activeRange.getRowIndex() - 1;
    let endRow = startRow + activeRange.getValues().length; // exclusive
    // create title
    for (let i = Math.max(firstTaskRow, startRow); i < Math.min(values.length, endRow); i++) {
      let session = values[i];
      session[eventFinalTitleCol] = createEventTitle_(session);
    }
    range.setValues(values);
  }
}

//Sets up the calendar with the events on the spreadsheet
function setUpCalendarTasks_(values, fullRange, selectedRange) {
  let cal = CalendarApp.getCalendarById(calendarId);

  let startRow = selectedRange.getRowIndex() - 1;
  let endRow = startRow + selectedRange.getValues().length; // exclusive

  // Start at 1 to skip the header row.
  for (let i = Math.max(1, startRow); i < Math.min(values.length, endRow); i++) {
    let session = values[i];
    let title = session[eventSubjectCol];
    let start = new Date(session[eventStartCol]);
    let end = new Date(session[eventEndCol]);
    let combinedDescription = session[eventDescriptionCol];
    if (session[eventDescriptionCol] != "" || session[eventAttachmentCol] != "")
      combinedDescription += "\n\n";
    combinedDescription += session[eventAttachmentCol];

    let options = {description: combinedDescription, location: session[eventLocationCol], sendInvites: false};
    let event = cal.createEvent(title, start, end, options)
        .setGuestsCanSeeGuests(true);
    session[eventIDCol] = event.getId();
  }
  fullRange.setValues(values); // TODO: figure out if this line can be removed!
}

//Helper function to make stuff formatted better for the calendar import
function joinDateAndTime_(date, time) {
  date = new Date(date);
  date.setHours(time.getHours());
  date.setMinutes(time.getMinutes());
  return date;
}

//Adds everyone on the roster to the event
function addAllGuests(eventId, values, cal){
  for (let i = 3; i < values.length; i++){
    let email = values[i][1];
    cal.getEventById(eventId).addGuest(email);
  }
}

/*****************************************************************************
 *****************************************************************************
 * B: RANDOMIZING SCHEDULES
 * This includes helper functions for managing schedule times,
 * as well as the functions that use these helpers. The aim is to create
 * random schedules, then assign to the staff. The script will avoid conflicts
 * amongst random events, but some OCs may be tasked to specific events, so
 * Tasking HOCs can manually fix those afterwards.
 * I recommend setting manual events first.
 *****************************************************************************
 *****************************************************************************/

/***********************************************************************
 * Helper function for finding the latest event in `scheduleTimes`
 * that begins strictly before `start`. Returns the index of that
 * event, or -1 if `start` is earlier than anything in the schedule.
 * 
 * `scheduleTimes`: a list of integer pairs, where each pair (x, y)
 *                  represents the start and ending unix timestamps
 *                  (in seconds) of a busy period.
 * `start`: start time, represented as a unix timestamp in seconds
 * 
 * Precondition: `scheduleTimes` is sorted by the start times in
 *               increasing order
 * Precondition: `scheduleTimes` contains no overlapping busy periods.
 ************************************************************************/
 function findLastPrecursorEvent_(scheduleTimes, start) {
  // 0) Empty schedule has no precursors
  //    (`start` is earliest event)
  if (scheduleTimes.length == 0)
    return -1;

  // 1) Binary search from [low, hi) to find latest event
  //    that begins strictly before `start`. At the end of
  //    the loop, `low` should be at the last element that
  //    we want to check
  let low = 0;
  let hi = scheduleTimes.length;
  let floorIndex = -1;

  while (low < hi - 1)
  {
    let mid = Math.floor((low + hi) / 2);
    let midVal = scheduleTimes[mid][0];

    // midEvent begins at or after `start`, so NOT a candidate
    // check an earlier event in the next loop 
    if (start <= midVal)
      hi = mid;
    else // midEvent is candidate, check a later event in the next loop
      low = mid;
  }
  // hi == lo + 1, meaning we've narrowed down to one event.

  // CORRECTNESS/SANITY CHECK:
  // if `low` > 0, then `low` had to have increased,
  // which is only possible if start > scheduleTimes[`low`][0].
  // 
  // The contrapositive means, if start <= scheduleTimes[`low`][0],
  // then `low` == 0 (`low` can only increase from 0), which means
  // start <= scheduleTimes[0][0], so it precedes the first event.
  if (start <= scheduleTimes[low][0])
    floorIndex = -1;
  else
    floorIndex = low;
  
  return floorIndex;
 }

/********************************************************************
 * Returns true if the interval [`start` - `buffer`, `end + buffer`]
 * is free in `schedule`, and returns false otherwise.
 * 
 * `scheduleTimes`: a list of integer pairs, where each pair (x, y)
 *                  represents the start and ending unix timestamps
 *                  (in seconds) of a busy period.
 * `start`: start time, represented as a unix timestamp in seconds
 * `end`: end time, represented as a unix timestamp in seconds
 * `buffer`: optional buffer time in seconds surrounding an event
 * 
 * Precondition: `scheduleTimes` is sorted by the start times in
 *               increasing order
 * Precondition: `scheduleTimes` contains no overlapping busy periods.
 * Precondition: `start` < `end`
 ********************************************************************/
function isScheduleFree_(scheduleTimes, start, end, buffer = 0) {
  // 0) Empty schedule is always free
  if (scheduleTimes.length == 0)
    return true;
  
  // 1) Find the lastest event that starts before `start`.
  let floorIndex = findLastPrecursorEvent_(scheduleTimes, start);

  // 2) Determine if (start, end) begins at the start, end, or middle of the schedule
  if (floorIndex <= -1) // `start` comes before the beginning of any event in the schedule
    return end + buffer < scheduleTimes[0][0]; // must end at least `buffer` seconds before the first event
  else if (floorIndex >= scheduleTimes.length - 1) // `start` comes after the beginning of any event in the schedule
    return start - buffer > scheduleTimes[scheduleTimes.length - 1][1]; // must start at least `buffer` seconds after the last event
  return start - buffer > scheduleTimes[floorIndex][1] && end + buffer < scheduleTimes[floorIndex + 1][0];
}

/*********************************************************************
 * Returns a new schedule containing the busy periods of
 * `scheduleTimes` as well as [`start`, `end`] in the correct space.
 * 
 * `schedule`: a list of integer pairs, where each pair (x, y)
 *             represents the start and ending unix timestamps (in
 *             seconds) a busy period.
 * `start`: start time, represented as a unix timestamp in seconds
 * `end`: end time, represented as a unix timestamp in seconds
 * 
 * Precondition: `scheduleTimes` is sorted by the start times in
 *               increasing order
 * Precondition: `scheduleTimes` contains no overlapping busy periods.
 * Precondition: `start` < `end`
 *******************************************************************/
 function addScheduleTime_(scheduleTimes, start, end) {
  // 0) Empty schedule is free, no need to split
  if (scheduleTimes.length == 0)
    return [[start, end]];

  // 1) Find the lastest event that starts before `start`
  let floorIndex = findLastPrecursorEvent_(scheduleTimes, start);

  // 2) Prepend or append if earliest or latest...
  if (floorIndex == -1)
    return [[start, end]].concat(scheduleTimes);
  else if (floorIndex == scheduleTimes.length - 1)
    return scheduleTimes.concat([[start, end]]);
  else
  {
    //  ...or split array into a1 and a2
    let a1 = scheduleTimes.slice(0, floorIndex+1);
    let a2 = scheduleTimes.slice(floorIndex+1, scheduleTimes.length);

    // 3) Append [start, end] to a1
    a1 = a1.concat([[start, end]]);

    // 4) Merge a1 and a2 and return
    return a1.concat(a2);
  }

 }

/*****************************************************************************************************
 * Each staff member will have a distinct combination of events (aka a schedule),
 * so this function creates those schedules to be assigned to staff.
 * 
 * `eventValues`: 2D array of the values from a spreadsheet (assumes first row is a header)
 * `staffCount`: the number of staff of a specific type (OC, OL, or HOC)
 * `staffColumn`: the column used to source the number of a staff type needed at an event
 * `scheduleTimes`: 2D array of time intervals (two-element lists) in unix seconds, with existing busy time intervals
 * Returns a tuple of (schedules, remainder, largest schedule size).
 * 
 * If a "specific" event needs x people, then it needs to appear in x combinations
 * So we determine the average amount of events an OC would need to be tasked to,
 * and that'll dictate the expected size of your combinations. Each event has a slot per staff member.
 * 
 * We check for conflicts using a parallel list of that represents schedules as lists of
 * time intervals (themselves being arrays of two elements because js doesn't have tuples oh my days).
 * These intervals correspond to when a staff member is busy, and are checked and updated accordingly.
 ******************************************************************************************************/
function createRandomSchedules_(eventValues, staffCount, staffColumn, scheduleTimes){
  // 0) Error checking
  if (staffCount <= 0)
  {
    Logger.log("Please have a positive staff count!");
    return [null, null, null];
  }

  // 1) Get event slots
  let allEventSlots = []; // a 2D list, where each list contains an event ID repeated for each slot
  let totalEventSlots = 0;
  for (let i = 1; i < eventValues.length; i++)
  {
    let thisEventRow = eventValues[i];
    let staffNeeded = thisEventRow[staffColumn];
    if (staffNeeded > 0)
    {
      let theseEventSlots = []; // Copy the event id for each slot
      for (let j = 0; j < staffNeeded; j++)
        theseEventSlots.push(thisEventRow);
      allEventSlots.push(theseEventSlots);
      totalEventSlots += staffNeeded;
    }
  }

  // 2) Determine number of events slots per staff member
  let eventsPerStaff = totalEventSlots/staffCount;

  // 3) Create schedules (and remainder list)
  // Do this via random event slot by random event slot, sequential schedule by sequential schedule
  // If no more events can be added (all schedules are full or any remaining events conflict with all schedules), then stop
  let schedules = [];
  // let scheduleTimes = []; // list of int pairs (sorted by start) for determining conflicts
  let remainder = [];
  let filledSchedules = new Set(); // skip these
  for (let i = 0; i < staffCount; i++)
  {
    schedules[i] = [];
  }
  let largestScheduleSize = 0;

  let currentSchedule = 0;
  while (allEventSlots.length > 0 && filledSchedules.size < staffCount)
  {
    if (!filledSchedules.has(currentSchedule)) // this schedule has not been *marked* as filled yet
    {
      // a) Is the schedule full via quantity?
      if (schedules[currentSchedule].length >= eventsPerStaff + 1)
      {
        filledSchedules.add(currentSchedule);
        // Logger.log("Schedule " + currentSchedule + " filled via quantity!");
      }
      else
      {
        let eventToMarking = new Map();
        let index = 0;
        let eventSlot = 0;

        let eventStart = 0;
        let eventEnd = 0;

        let valid = false;

        // b) Event validity:
        //    i) prevent dupes:
        //       Mark the index once you've seen it.
        //       If the current schedule contains this event
        //       and there are still event slots that haven't been marked, try again 
        //   ii) prevent conflicts:
        //       Get the start and end times (unix timestamp in seconds)
        //       and use the schedule free checker function.
        do
        {
          index = getRandomInt(0, allEventSlots.length);
          eventSlot = allEventSlots[index][0];
          eventToMarking.set(index, true);

          eventStart = new Date(eventSlot[eventStartCol]).getTime();
          eventEnd = new Date(eventSlot[eventEndCol]).getTime();

          // a valid event is not a duplicate and doesn't conflict with the schedule
          valid = schedules[currentSchedule].indexOf(eventSlot) <= -1 &&
                             isScheduleFree_(scheduleTimes[currentSchedule], eventStart, eventEnd, scheduleBuffer);
        } while (eventToMarking.size < allEventSlots.length && !valid);

        // c) Determine if allowed
        if (valid)
        {
          // d) add event ID and time
          schedules[currentSchedule].push(allEventSlots[index].pop());
          scheduleTimes[currentSchedule] = addScheduleTime_(scheduleTimes[currentSchedule], eventStart, eventEnd);

          // e) get rid of events that have all slots taken
          if (allEventSlots[index].length == 0)
          {
            let newEventSlots = [];
            for (let i = 0; i < allEventSlots.length; i++)
              if (i != index)
                newEventSlots.push(allEventSlots[i]);
            allEventSlots = newEventSlots;
          }

          // f) calculate largest schedule
          if (schedules[currentSchedule].length > largestScheduleSize)
            largestScheduleSize = schedules[currentSchedule].length;
        }
        else // Could not find valid event to add
        {
          filledSchedules.add(currentSchedule);
          // Logger.log("Schedule " + currentSchedule + " filled due to conflicts!");
        }
      }
    }
    // Logger.log("Schedule " + currentSchedule + ": " + schedules[currentSchedule]);
    // Logger.log(allEventIDs.length + " events left and " + filledSchedules.length + " filled schedules.");
    currentSchedule = (currentSchedule + 1) % staffCount;
  }

  // 4) Add remaining events to `remainder` by flattening allEventIDs
  // TODO: Convert to for loop to (slightly) improve performance
  remainder = allEventSlots.length > 0 ? allEventSlots.reduce((a, b) => a.concat(b)) : [];

  // 5) Return schedules and remaining events to assign
  return [schedules, remainder, largestScheduleSize];
}

/*************************************
 * Converts an event row into a title,
 * formatted as:
 * "Subject (Date) [ID]"
 *************************************/
function createEventTitle_(event) {
  return event[eventSubjectCol] + " (" + new Date(event[eventStartCol]).getTime() + " - " +
                                       new Date(event[eventEndCol]).getTime() + ") " +
                                  "[" + event[eventIDCol] + "]"; 
}
function scheduleToTitles_(schedule){
  return schedule.map(function(value) { return createEventTitle_(value)});
}
/*****************************************
 * Converts a schedule of events to just
 * a particular column
 *****************************************/
function scheduleToColumn_(schedule, column){
  return schedule.map(function(value) { return value[column] });
}

/*******************************************************
 * Returns a new schedule with padded entries.
 * You can optionally choose what to pad the space with.
 *******************************************************/
function padScheduleRow_(schedule, size, padding = ""){
  let padded = [].concat(schedule);
  let spaces = size - schedule.length;
  for (let i = 0; i < spaces; i++)
    padded.push(padding);
  return padded;
}

function TEST_writeToSheet(){
  let ss = SpreadsheetApp.getActive();
  let assignmentSheet = ss.getSheetByName('Sandbox(Test)');
  assignmentSheet.getRange("C3:F6").setValues([["\"hey", "baby", "how's", "it"],
                                               ["going?", "", "This", "beat"],
                                               ["is", "non", "-", "stop!\""],
                                               ["(Play", "Rhythm", "Heaven,", "please)"]]);
}

function TEST_countRowSizes() {
  let ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName("Sandbox(Test)2");
  let values = sheet.getDataRange().getValues();

  for (let i = 0; i < values.length; i++) {
    Logger.log(values[i].length);
  }
}

/*********************************************************************
 * Helper function to create and assign schedules for a position type.
 * `posSlotCol`: the column in `eventSheetName` that corresponds to
 * the amount of `position` needed.
 *********************************************************************/
function assignSchedulesForPosition_(position, posSlotCol, taskValues, assignmentSheetName)
{
  // (1) Get information about roster + role
  let ss = SpreadsheetApp.getActive();
  let assignmentSheet = ss.getSheetByName(assignmentSheetName);
  let assignmentValues = assignmentSheet.getDataRange().getValues();
  let rosterSheet = ss.getSheetByName('Roster');
  let rosterPositions = rosterSheet.getRange(rosterPosLetter + hocMin + ":" +
                                             rosterPosLetter + (ocMax+1)).getValues(); // 2D array of the values themselves
  let staffCount = rosterPositions.filter(function(value) { return value == position} ).length;

  // (2) Create scheduleTimes based on manual tasks
  ///    as an argument by reading from assignmentSheet 
  let scheduleTimes = [];
  let manualCount = [];
  for (let staff = 0; staff < staffCount; staff++)
  {
    scheduleTimes[staff] = [];
    let staffSchedule = assignmentValues[staff+1];
    // Logger.log(staffSchedule);
    manualCount[staff] = 0;
    for (let j = assignmentStartCol; j < staffSchedule.length; j++)
    {
      let event = staffSchedule[j];
      let time = eventTimeParse_(event);
      // Logger.log(time);
      if (time) {
        scheduleTimes[staff] = addScheduleTime_(scheduleTimes[staff], time[0], time[1])
        manualCount[staff]++;
      }
    }
    // Logger.log(scheduleTimes[staff]);
    // Logger.log(scheduleTimes[staff]);
  }

  // (3) Create schedule by counting the number of people in the position
  let assignments = createRandomSchedules_(taskValues, staffCount, posSlotCol, scheduleTimes);
  let schedules = assignments[0];
  let remainder = assignments[1];
  let maxEvents = assignments[2];


  let lastColMain = String.fromCharCode(assignmentStartLetter.charCodeAt(0) + (maxEvents - 1));
  let lastColRemain = String.fromCharCode(assignmentStartLetter.charCodeAt(0) + (remainder.length - 1));

  // (4) Write to event spreadsheet (map via create event title, then pad using the max length)
  // This is for us to manually fix overlaps
  let scheduleTitles = schedules.map(function(value) { return scheduleToTitles_(value) });
  scheduleTitles = scheduleTitles.map(function(value) { return padScheduleRow_(value, maxEvents) });
  
  // Set value row by row
  // assignmentSheet.getRange(assignmentStartLetter + assignmentStartNum + ":" + (lastColMain) + (assignmentStartNum + staffCount - 1)).setValues(scheduleTitles);
  for (let staff = 0; staff < staffCount; staff++)
  {
    let staffSchedule = schedules[staff];
    let firstRandomCol = assignmentStartCol + manualCount[staff];
    for (let event = 0; event < staffSchedule.length; event++)
      assignmentValues[staff+1][firstRandomCol + event] = createEventTitle_(staffSchedule[event]);
  }
  assignmentSheet.getDataRange().setValues(assignmentValues);
  

  if (remainder.length > 0)
  {
    // Write remainder
    let remainderTitles = scheduleToTitles_(remainder);
    assignmentSheet.getRange(assignmentStartLetter + (assignmentStartNum + staffCount) + ":"
                        + (lastColRemain) + (assignmentStartNum + staffCount)).setValues([remainderTitles]);
  }

  // // Write to ID spreadsheet (for invites)
  // let scheduleIDs = schedules.map(function(value) { return scheduleToColumn(value, idColumn) });
  // let idSheet = ss.getSheetByName(idSheetName);
  // idSheet.getRange("C2:I" +(1+count)).setValues(scheduleIDs);
}

/******************************************************
 * Assigns random schedules for OCs, OLs, and HOCs
 * And writes them to a spreadsheet by concactenating
 * the subject, date, start time, end time, and id.
 ******************************************************/
function assignRandomSchedules()
{
  // 0) Setup: Reading from Roster and SpecificTasks, Writing to SpecificAssignments
  let ss = SpreadsheetApp.getActive();
  let taskSheet = ss.getSheetByName('RandomTasks');
  let taskValues = taskSheet.getDataRange().getValues(); // 2D array of the values themselves

  // 1) Create OC schedules, write to spreadsheet
  assignSchedulesForPosition_("OC", ocsNeededCol, taskValues, "SpecificAssignments(OC)");

  // 2) Create OL schedules, write to spreadsheet
  assignSchedulesForPosition_("OL", olsNeededCol, taskValues, "SpecificAssignments(OL)");

  // 3) Create HOC schedules, write to spreadsheet
  //    (probably unnecessary tbh since there's ~eight of them, and will either be
  //     at their taskings' events or can just decide amongst each other who does what)
  assignSchedulesForPosition_("HOC", hocsNeededCol, taskValues, "SpecificAssignments(HOC)");
}

/*****************************************************************************
 *****************************************************************************
 * C: INVITING TO BIG EVENTS
 * These are events that EVERY staff member is invited to (e.g. "tent events")
 *****************************************************************************
 *****************************************************************************/

function sendBigInvites() {
  sendMassInvites_('BigTasks');
}

function sendTrainingInvites() {
  sendMassInvites_('TrainingSchedules');
}

/*************************************************************
 * Send Mass Invites to Selected Events from `eventSheetName`
 *************************************************************/
// Based on code from `addAllGuests` function, assumes big events have been created
function sendMassInvites_(eventSheetName) {
  let cal = CalendarApp.getCalendarById(calendarId);

  let ss = SpreadsheetApp.getActive();
  let rosterSheet = ss.getSheetByName('Roster');
  let rosterValues = rosterSheet.getDataRange().getValues();
  let bigTaskSheet = ss.getSheetByName(eventSheetName);
  let bigTaskValues = bigTaskSheet.getDataRange().getValues();

  let activeRange = bigTaskSheet.getActiveRange();

  if (activeRange) {
    let startRow = activeRange.getRowIndex() - 1;
    let endRow = startRow + activeRange.getValues().length; // exclusive

    // [start, end)
    for (let event = Math.max(1, startRow); event < Math.min(bigTaskValues.length, endRow); event++) {
      // invite each role (we don't necessarily assume each role is contiguous on the roster here,
      //                   although there's no reason they wouldn't be if you follow the formatting guidelines...)
      for (let staff = hocMin - 1; staff < hocMax - 1; staff++) {
        let email = rosterValues[staff][rosterEmailCol];
        if (indexOfGuestEmailList_(cal.getEventById(bigTaskValues[event][eventIDCol]).getGuestList(), email) < 0)
          cal.getEventById(bigTaskValues[event][eventIDCol]).addGuest(email);
      }
      for (let staff = olMin - 1; staff < olMax - 1; staff++) {
        let email = rosterValues[staff][rosterEmailCol];
        if (indexOfGuestEmailList_(cal.getEventById(bigTaskValues[event][eventIDCol]).getGuestList(), email) < 0)
          cal.getEventById(bigTaskValues[event][eventIDCol]).addGuest(email);
      }
      for (let staff = ocMin - 1; staff < ocMax - 1; staff++) {
        let email = rosterValues[staff][rosterEmailCol];
        if (indexOfGuestEmailList_(cal.getEventById(bigTaskValues[event][eventIDCol]).getGuestList(), email) < 0)
          cal.getEventById(bigTaskValues[event][eventIDCol]).addGuest(email);
      }
    }
  }
  else
  {
    console.log("Please select a range in the " + "\"" + eventSheetName + "\"!");
  }
}

/*****************************************************************************
 *****************************************************************************
 * D: INVITING TO SPECIFIC (RANDOM + MANUAL)
 * Random schedules are created with the eventID in the "title", in case
 * any manual changes have to be made. Rather than dedicating a whole nother
 * sheet to the IDs, we can just parse the event ID from the title and invite.
 *****************************************************************************
 *****************************************************************************/

function sendSelectedOCsInvites() {
  sendSelectedStaffInvites_('OC');
}

function sendSelectedOLsInvites() {
  sendSelectedStaffInvites_('OL');
}

function sendSelectedHOCsInvites() {
  sendSelectedStaffInvites_('HOC');
}

function sendSelectedStaffInvites_(role) {
  let ss = SpreadsheetApp.getActive();
  let assignmentSheet = ss.getSheetByName('SpecificAssignments(' + role + ')');

  let activeRange = assignmentSheet.getActiveRange();

  if (activeRange) {
    sendSpecificInvites_(0, assignmentSheet, activeRange);
  }
  else {
    console.log("Please select a continuous range in the " + "\"SpecificAssignment(" + role + ")\"!");
  }
}

function sendSpecificInvites_(count, assignmentSheet, selectedRange) {
  let cal = CalendarApp.getCalendarById(calendarId);
  let assignmentValues = assignmentSheet.getDataRange().getValues();

  if (selectedRange) // trying to invite specific OCs (essentially, updating specific schedules)
  {
    let startRow = selectedRange.getRowIndex() - 1;
    let endRow = startRow + selectedRange.getValues().length; // exclusive

    for (let staff = startRow; staff < endRow; staff++) {
      let email = assignmentValues[staff][assignmentEmailCol];
      for (let event = assignmentStartCol; event < assignmentValues[staff].length; event++) {
        let eventID;
        if ((eventID = eventIDParse_(assignmentValues[staff][event])))
        {
          let calEvent = cal.getEventById(eventID);
          if (indexOfGuestEmailList_(calEvent.getGuestList(), email) < 0)
            cal.getEventById(eventID).addGuest(email);
        }
      }
    }
  }
  else { // Mass invite
    for (let staff = 1; staff < count + 1; staff++) {
      for (let event = assignmentStartCol; event < assignmentValues[staff].length; event++) {
        let email = assignmentValues[staff][assignmentEmailCol];
        let eventID;
        if ((eventID = eventIDParse_(assignmentValues[staff][event])))
          cal.getEventById(eventID).addGuest(email);
      }
    }
  }
}

//Didn't end up using it but leaving it in for future use just in case
function sendInvites_(user, response) {
  // let id = ScriptProperties.getProperty('calId');
  let cal = CalendarApp.getCalendarById(calendarId);
  for (let i = 0; i < response.length; i++) {
    cal.getEventSeriesById(response[i][5]).addGuest(user.email);
  }
}

//Sets up the staff for events in the sheet
// function setUpGuests(){
//   let ss = SpreadsheetApp.getActive();
//   let sheet1 = ss.getSheetByName('Sheet15'); //will have to change this if you want to do other sheets
//   let range1 = sheet1.getDataRange();
//   let values1 = range1.getValues();
//   setStaff(values1, range1);
//   //setRandomStaff(values1, range1)
// }

/*****************************************************************************
 *****************************************************************************
 * E: PARSER HELPERS
 *****************************************************************************
 *****************************************************************************/

 /*******************************************
 * Parses the event ID from the event title,
 * as created by createEventTitle_() via
 * regex pattern matching.
 * Returns the ID.
 ******************************************/
function eventIDParse_(title) {
  let idPattern = /\[[A-Za-z0-9]+@google.com\]/g;
  let re = idPattern.exec(title);
  if (re == null)
  {
    // Logger.log("Regex failed on title \"" + title + "\"! Perhaps you didn't use \`eventToTitle\` on it?");
    return null;
  }

  let id = title.substring(re.index + 1, title.length-1);
  // Logger.log(id);
  return id;
}

 /***************************************************
 * Parses the event Time RAnge from the event title,
 * as created by createEventTitle_() via
 * regex pattern matching.
 * Returns the start and end time.
 ***************************************************/
function eventTimeParse_(event) {
  let idPattern = /\([0-9]+ - [0-9]+\)/g;
  let re = idPattern.exec(event);
  if (re == null)
  {
    // Logger.log("Regex failed on title \"" + title + "\"! Perhaps you didn't use \`eventToTitle\` on it?");
    return null;
  }
  let range = re[0];
  let hyphenIdx = range.indexOf('-');

  let start = parseInt(range.substring(1, hyphenIdx - 1));
  let end = parseInt(range.substring(hyphenIdx + 2, range.length));

  return [start, end];
}

 /******************************
 * Returns a list of event IDs.
 *******************************/
function eventIDParseRow_(titles) {
  return titles.map(function(value) { return eventIDParse_(value); });
}

/*****************************************************************************
 *****************************************************************************
 * F: MISCELLANEOUS HELPERS
 *****************************************************************************
 *****************************************************************************/
//Allows you to delete all the events for a specific range (we used this to erase things that we messed up on)
function delete_events()
{
  // EDIT THESE VALUES BEFORE UNCOMMENTING
  // var fromDate = new Date(2022,7,10);
  // var toDate = new Date(2022,7,31);
  var calendar = CalendarApp.getCalendarById(calendarId);

  var events = calendar.getEvents(fromDate, toDate);
  for(var i=events.length-1; i>=0; i--)
    events[i].deleteEvent();
}

function indexOfGuestEmailList_(guestList, email) {
  for (let i = 0; i < guestList.length; i++)
  {
    if (guestList[i].getEmail() == email)
      return i;
  }
  return -1;
}


/*****************************************************************************
 *****************************************************************************
 * G: OTHER FUNCTIONS FROM ELLIA'S SCRIPT
 *****************************************************************************
 *****************************************************************************/
//My original function for pulling staff from spreadsheet
//Didn't end up using this exact function but leaving it here in case there's anything other people want to use
// function getStaff_(ocNum, olNum, hocNum){
//   result = '';
//   first = true;
//   for (let i = 0; i < ocNum; i++){
//     if (first){
//       result += ocStart;
//       first = false;
//     }
//     else{
//       result += ',' + ocStart;
//     }
//     ocStart = (ocStart + 1)%125;
    
//   }
//   for (let i = 0; i < olNum; i++){
//     if (first){
//       result += olStart;
//       first = false;
//     }
//     else{
//       result += ',' + olStart;
//     }
//     olStart = (olStart + 1)%28;
    
//   }
//   for (let i = 0; i < hocNum; i++){
//     if (first){
//       result += hocStart;
//       first = false;
//     }
//     else{
//       result += ',' + hocStart;
//     }
//     hocStart = (hocStart + 1)%7;
    
//   }
//   return result;
// }

// //This was for finding their numbers based on their names I think, but also didn't end up using this in my final implementation
// function findStaffNum_(names){
//   let ss = SpreadsheetApp.getActive();
//   let sheet = ss.getSheetByName('Roster');
//   let range = sheet.getDataRange();
//   namesList = names.split(',');
//   result = ""
//   for (let name of namesList){
//     result += "," + MATCH(name, range, 0).toString();
//   }
//   return result;
// }

// //For generating random staff numbers, also was in original implementation
// function getRandomInt_(min, max) {
//   min = Math.ceil(min);
//   max = Math.floor(max);
//   return Math.floor(Math.random() * (max - min) + min); 
// }
