/**
 * @file Meeting Tagger Tool v2.9.13
 * @description MAJOR UPDATE: AUTOTAG !NOT FILTERS
 * /**
 * KEYWORD MATCHING LOGIC (v2.9.0)
 * ------------------------------
 * Supports "Smart Exclusions" using the syntax: keyword(!exclude1|exclude2)
 *
 * HOW IT WORKS (Single Rule Logic):
 * A single rule handles both the "Match" and the "Exclusion". You do not need
 * separate rules for the same keyword.
 *
 * LOGIC FLOW:
 * 1. Search for 'keyword'. If not found -> NO MATCH.
 * 2. If found, check for 'exclusions'.
 * - If an exclusion word is found -> NO MATCH (The rule kills itself).
 * - If no exclusion is found      -> MATCH (Success).
 *
 * EXAMPLE: Rule = "vce(!coverage)"
 * - Input: "VCE Salt Lake City"         -> MATCH (Keyword 'vce' found, 'coverage' absent).
 * - Input: "Your Name - VCE - Coverage" -> NO MATCH (Keyword found, but blocked by 'coverage').
 *
 * v2.9.1 --- added Discrepancy Check Table to see Calendar vs Radar Events 
 * v2.9.2 --- added Instructions Tab
 * v2.9.3 --- added Hidden RADAR sheet as well as other 2 RADAR Sheets
 * v2.9.4 --- added gCal Labels in Tags Sheet
 *    --- added strict REGEX so that 'trial doesn't match on words like 'industrial'
 * v2.9.5 --- backup existing Filter Lists and Tags when 'repairing' a sheet
 * v2.9.6 --- added user-defined LocationExclusions for removing suggestions of 'InPerson' on location keywords
 * v2.9.7 --- added Labels support for custom names. Added Visuals Barchart and APP_CONFIG brige Flags to disable/enable it
 * v2.9.8 --- added Manager Rollup Table to Visuals Dashboard and associated Bridge Script feature flags/support
 * v2.9.9 --- clean up and addition of commonly used suggestions
 * v2.9.10 -- separated 'manager tables' to a new tab 'Team Visuals' for clarity to begin rework of IC Visuals with AE breakdowns
 * v2.9.11 -- added AE Time Breakdown Table to Visuals and Removed 'Labels' Table, also improved REGEX logic
 * v2.9.12 -- bug fix for 'MinutesInterval' parameter as well as errors regarding empty 'Filter Lists'
 *    --- bug fix for meetings parsing when broken html links are appended to meeting tags
 * v2.9.13 -- added scaffolding and logic (region 6a) for Manager Dashboard and Team Rollup Metrics using API (not RADAR)
 *    --- bug fix around 'tag validation' function and updated aeHeaderRenameMap with additional tag types
 *    --- bug fix regarding NAVAN events not having accepted invitees causing some filters to skip NAVAN events
 */

// =================================================================
// #region 1. SPREADSHEET UI & MENU CREATION (v3.0.0 - Feature Flags)
// =================================================================

// NON BRIDGE FUNCTION 
function onOpen() {
  // Pass the global APP_CONFIG to the menu creator
  createMeetingToolsMenu(APP_CONFIG);
}

function createMeetingToolsMenu(options = {}) {
  // Default: Radar is ENABLED unless explicitly set to false
  const enableRadar = options.features?.radar !== false; 
  
  // Check if the team manager dashboard is enabled
  const enableTeam = options.team?.enabled === true; 

  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Meeting Tools');
  
  // Standard Items
  menu.addItem('Save Config and Refresh Untagged List', 'findEventsMissingTag')
      .addSeparator()
      .addItem('Apply Changes to Calendar', 'applyBatchChanges')
      .addSeparator()
      .addItem('Generate Visuals Dashboard', 'generateDashboard')
      .addSeparator();

  // MANAGER CONDITIONAL: Only show if options.team.enabled is true
  if (enableTeam) {
    menu.addItem('Generate Manager Dashboard', 'runManagerDashboard')
      .addSeparator();
  }

  // Config Tools
  menu.addItem('Config: Update Quarter List', 'updateQuarterDropdown')
      .addItem('Config: Update Tag Dropdowns & Colors', 'updateTagDropdownsAndColors')
      .addItem('Config: Validate Tag Keywords', 'validateTagKeywords')
      .addSeparator();

  // Admin Submenu (Built step-by-step to allow conditionals)
  const adminMenu = ui.createMenu('Admin: Initialize/Repair Sheets');
  
  adminMenu.addItem('Initialize ALL Sheets (Full Reset)', 'initializeAllSheets')
           .addSeparator()
           .addItem('Repair "Instructions" Sheet Only', 'repairInstructionsSheet')
           .addItem('Repair "Untagged Meetings" Sheet Only', 'repairUntaggedSheet')
           .addItem('Repair "Config" Sheet Only', 'repairConfigSheet')
           .addItem('Repair "Filter Lists" Sheet Only', 'repairFilterSheet')
           .addItem('Repair "Tags" Sheet Only', 'repairTagsSheet');

  // CONDITIONAL RADAR MENU ITEMS
  if (enableRadar) {
      adminMenu.addSeparator()
               .addItem('Repair "SE Radar Meetings" Sheet Only', 'repairRadarMainSheet')
               .addItem('Repair "All SE Radar Meetings CFQ" Sheet Only', 'repairRadarFilteredSheet') 
               .addItem('Repair "SFDC Radar Master" Sheet Only', 'repairRadarImportSheet');
  }

  // Attach Submenu and Build
  menu.addSubMenu(adminMenu);
  menu.addToUi();
}
// #endregion

// =================================================================
// #region 1a. Google Calendar Event Lables Table (CONSTANT) - (v2.9.4)
// =================================================================

// Google Calendar Color ID Mapping
const EVENT_COLOR_MAP = {
  "Lavender": "1",
  "Sage": "2",
  "Grape": "3",
  "Flamingo": "4",
  "Banana": "5",
  "Tangerine": "6",
  "Peacock": "7",
  "Graphite": "8",
  "Blueberry": "9",
  "Basil": "10",
  "Tomato": "11",
  "Default": "" 
};

// #endregion

// =================================================================
// #region 2. SCRIPT ENTRY POINTS (v3.5.1 - Fix Zero Interval)
// =================================================================

/**
 * MASTER TRIGGER: Handles all simple edits.
 * Routes traffic to the correct helper function based on the sheet name.
 */
function onEdit(e) {
  const sheet = e.range.getSheet();
  const name = sheet.getName();

  if (name === "Untagged Meetings") {
    // Handle coloring edits on the main list
    updateMeetingTag(e, true); 
  } 
  else if (name === "Tags") {
    // Handle "Changes Detected" flag on Tags sheet
    onEditTagsSheet(e);
  }
}

// LIBRARY FUNCTION: Executes the manager dashboard pipeline
function executeManagerDashboard(options) {
  const ui = SpreadsheetApp.getUi();
  if (!options || !options.team || !options.team.enabled) {
      ui.alert("Manager Dashboard is disabled in APP_CONFIG.");
      return;
  }
  
  refreshTeamCalendarDatabase(options);
  generateManagerDashboard(options);
}

/**
 * Menu action to apply batch changes.
 */
function applyBatchChanges() {
  updateMeetingTag(null, false); 
}

// #endregion

// =================================================================
// #region 3. CORE LOGIC FUNCTIONS (v4.2.0 - URL & Token Safety)
// =================================================================

function findEventsMissingTag(forcedUserEmail = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = ss; 
  try {
    // --- 0. AUTO-REFRESH TRIGGER MANAGEMENT ---
    // This safely builds your timer in the background using Google's strict intervals
    try {
        const minutesRaw = parseInt(getConfig("MinutesInterval"), 10);
        let validMinutes = 0;
        
        // Google Apps Script strictly allows only 1, 5, 10, 15, or 30 minute intervals.
        if (!isNaN(minutesRaw) && minutesRaw > 0) {
            if (minutesRaw <= 1) validMinutes = 1;
            else if (minutesRaw <= 5) validMinutes = 5;
            else if (minutesRaw <= 10) validMinutes = 10;
            else if (minutesRaw <= 15) validMinutes = 15;
            else validMinutes = 30; // Caps at 30 to prevent API timeouts
        }

        const props = PropertiesService.getDocumentProperties();
        const activeInterval = props.getProperty('AutoRefreshInterval');

        // Only delete and recreate if the interval has actually changed (Saves API Quota)
        if (String(validMinutes) !== activeInterval) {
            const triggers = ScriptApp.getProjectTriggers();
            triggers.forEach(trigger => {
                if (trigger.getHandlerFunction() === 'findEventsMissingTag') {
                    ScriptApp.deleteTrigger(trigger);
                }
            });

            if (validMinutes > 0) {
                ScriptApp.newTrigger('findEventsMissingTag').timeBased().everyMinutes(validMinutes).create();
            }
            props.setProperty('AutoRefreshInterval', String(validMinutes));
        }
    } catch (triggerErr) {
        console.log("Trigger setup skipped: " + triggerErr.message);
    }

    // --- 1. CONFIGURATION ---
    const daysBack = parseInt(getConfig("DaysBack"), 10);
    const daysAhead = parseInt(getConfig("DaysAhead"), 10);
    const selectedQuarter = getConfig("QuarterOverride");
    const MAX_PATCH_DAYS_PAST = 500; 

    const ignorePhrasesRules = _parseMatchRule(_getFilterList("IgnorePhrases").join(","));
    const ignoreExactList = _getFilterList("IgnoreExactTitles");
    const ignoreFromEmailsList = _getFilterList("IgnoreFromEmails").map(s => s.toLowerCase());
    const ignoreToEmailsList = _getFilterList("IgnoreToEmails").map(s => s.toLowerCase());

    const useIgnorePhrases = (getConfig("IgnorePhrases") || "false").toString().toLowerCase() === "true";
    const useIgnoreExact = (getConfig("IgnoreExactTitles") || "false").toString().toLowerCase() === "true";
    const useIgnoreFrom = (getConfig("IgnoreFromEmails") || "false").toString().toLowerCase() === "true";
    const useIgnoreTo = (getConfig("IgnoreToEmails") || "false").toString().toLowerCase() === "true";
    const useAutoSeLead = (getConfig("AutoSELeadKeywords") || "false").toString().toLowerCase() === "true";
    const useAutoSeCover = (getConfig("AutoSECoverageKeywords") || "false").toString().toLowerCase() === "true"; 
    const useAutoInPerson = (getConfig("AutoInPersonKeywords") || "false").toString().toLowerCase() === "true";

    const includeAllDayEvents = (getConfig("IncludeAllDayEvents") || "false").toString().toLowerCase() === "true";
    const externalFilter = (getConfig("ExternalAttendees") || "false").toString().toLowerCase() === "true";
    const filterForMissingSETags = (getConfig("FilterForMissingSETags") || "false").toString().toLowerCase() === "true"; 
    const showTaggedMeetings = (getConfig("ShowTaggedMeetings") || "false").toString().toLowerCase() === "true";
    const locationFilterEnabled = (getConfig("LocationSetsInPerson") || "false").toString().toLowerCase() === "true";
    const locationRequiresExternal = (getConfig("LocationRequiresExternal") || "false").toString().toLowerCase() === "true";
    
    const modifier = getConfig("InPersonModifier") || "";
    const seLeadTagText = getConfig("SELeadTagText") || "";
    const seCoverTagText = getConfig("SECoverageTagText") || "";
    
    const tagPrefix = "Verkada Meeting Tag: ";
    const attendanceFilter = getConfig("AttendanceFilter");
    
    const virtualExclusions = ["teams", "zoom", "webex", "online", "call", "gmeet", "meet.google"];
    const locationExclusions = _getFilterList("LocationExclusions").map(s => s.toLowerCase());

    // --- 2. PREPARE TAG DATA ---
    const tagsSheet = file.getSheetByName("Tags");
    const tagsData = tagsSheet.getRange("B2:G" + tagsSheet.getLastRow()).getValues();
    
    const tagWaterfall = [];
    const tagMap = new Map(); 

    tagsData.forEach(row => {
      const tagName = row[1];
      const abbrTag = row[2] ? String(row[2]).replace(/;/g, "").trim() : "";
      const keywordString = row[3];
      
      const checkTitle = (row[4] === true || String(row[4]).toUpperCase() === 'TRUE');
      const checkDesc = (row[5] === true || String(row[5]).toUpperCase() === 'TRUE');

      if (abbrTag) {
        tagMap.set(abbrTag, tagName); 
        if (keywordString && (checkTitle || checkDesc)) {
             const rules = _parseMatchRule(keywordString);
             tagWaterfall.push({
                 tagName: tagName, 
                 abbrTag: abbrTag, 
                 rules: rules,
                 checkTitle: checkTitle,
                 checkDesc: checkDesc
             });
        }
      }
    });

    const autoSeLeadRules = useAutoSeLead ? _parseMatchRule(_getFilterList("AutoSELeadKeywords").join(",")) : [];
    const autoSeCoverRules = useAutoSeCover ? _parseMatchRule(_getFilterList("AutoSECoverageKeywords").join(",")) : [];
    const autoInPersonRules = useAutoInPerson ? _parseMatchRule(_getFilterList("AutoInPersonKeywords").join(",")) : [];
    const nightsAwayRules = _parseMatchRule(_getFilterList("NightsAwayKeywords").join(","));

    // --- 3. DATE LOGIC ---
    let start, end;
    if (selectedQuarter && selectedQuarter !== "" && selectedQuarter !== "None") {
        const parts = selectedQuarter.replace(" (Current)", "").trim().match(/FY(\d{2})\s*Q(\d)/);
        if (!parts) throw new Error(`Invalid quarter format: "${selectedQuarter}".`);
        const dates = _getDatesForVerkadaFQ(parseInt(parts[1], 10) + 2000, parseInt(parts[2], 10));
        start = dates.start; end = dates.end;
    } else {
        start = new Date(); start.setHours(0, 0, 0, 0); start.setDate(start.getDate() - daysBack);
        end = new Date(); end.setHours(23, 59, 59, 999); end.setDate(end.getDate() + daysAhead);
    }
    
    // --- 4. FETCH EVENTS ---
    let rawEmail = forcedUserEmail;
    
    // Fallback 1: Active User (When you click the button)
    if (!rawEmail) {
        try { rawEmail = Session.getActiveUser().getEmail(); } catch (e) { rawEmail = ""; }
    }
    // Fallback 2: Effective User (When the background timer runs it for you)
    if (!rawEmail) {
        try { rawEmail = Session.getEffectiveUser().getEmail(); } catch (e) { rawEmail = ""; }
    }
    
    const userEmail = (rawEmail || "").toString().toLowerCase();

    let calendarId;
    try {
        const cal = CalendarApp.getCalendarById(userEmail);
        if (!cal) throw new Error("Calendar not found");
        calendarId = cal.getId();
    } catch (e) {
        calendarId = CalendarApp.getDefaultCalendar().getId();
    }

    let allEvents = [], pageToken = null;
    do {
      const response = Calendar.Events.list(calendarId, { timeMin: start.toISOString(), timeMax: end.toISOString(), singleEvents: true, orderBy: 'startTime', maxResults: 2500, pageToken: pageToken });
      if (response.items) allEvents = allEvents.concat(response.items);
      pageToken = response.nextPageToken;
    } while (pageToken);

    // --- 5. PROCESS EVENTS ---
    const events = allEvents;
    const noTagMeetings = [], aeTaggedMeetings = [], fullyTaggedMeetings = [];

    for (const event of events) {
        if (event.eventType === 'outOfOffice') continue; 
        if (event.eventType === 'workingLocation') continue;

        const eventStartTimeMs = new Date(event.start.dateTime || event.start.date).getTime();
        const todayMs = new Date().getTime();
        if ((todayMs - eventStartTimeMs) > (MAX_PATCH_DAYS_PAST * 24 * 60 * 60 * 1000)) {
            continue; 
        }

        let durationInMinutes = 0;
        let nightsCount = 0;
        
        const titleRaw = event.summary || "";
        const titleSearch = titleRaw.replace(/(https?:\/\/[^\s]+)/g, '').replace(/\S{50,}/g, '').toLowerCase();
        let descRaw = (event.description || "").replace(/<br\s*\/?>/gi, '\n').replace(/<p>/gi, '\n').replace(/<\/p>/gi, '\n').replace(/<(?!(?:https?|tel):)[^>]*>/gi, ''); 
        const descSearch = descRaw.replace(/(https?:\/\/[^\s]+)/g, '').replace(/\S{50,}/g, '').toLowerCase();
        const fullSearch = titleSearch + " " + descSearch;

        const eventColorId = event.colorId || "";

        // NEW: Calculate isTrip up here so the RSVP filter can see it
        const isTrip = _isMatch(fullSearch, nightsAwayRules);

        if (event.start.date) {
            if (!includeAllDayEvents) continue; 
            
            // Allow bots (like Navan) to give you nights away credit
            if (isTrip) {
                const s = new Date(event.start.date);
                const e = new Date(event.end.date);
                const diff = Math.abs(e - s);
                const days = Math.ceil(diff / (1000 * 60 * 60 * 24));
                nightsCount = Math.max(0, days - 1); 
            }
        } else {
            const s = new Date(event.start.dateTime).getTime();
            const e = new Date(event.end.dateTime).getTime();
            durationInMinutes = Math.round((e - s) / 60000);
        }

        if (!titleRaw) continue;
        
        if (useIgnoreExact && ignoreExactList.includes(titleRaw)) continue;
        if (useIgnorePhrases && _isMatch(titleSearch, ignorePhrasesRules)) continue;
        
        const organizerEmail = event.organizer ? event.organizer.email.toLowerCase() : (event.creator ? event.creator.email.toLowerCase() : '');
        const attendees = (event.attendees || []).map(a => a.email.toLowerCase()).filter(Boolean);

        if (useIgnoreFrom && ignoreFromEmailsList.includes(organizerEmail)) continue;
        if (useIgnoreTo && attendees.some(a => ignoreToEmailsList.includes(a))) continue;
        
        const allParticipants = [...new Set([organizerEmail, ...attendees])];
        const isInternalEntity = email => email && (email.toLowerCase().endsWith("@verkada.com") || email.toLowerCase().endsWith("@resource.calendar.google.com") || email.toLowerCase().endsWith("@group.calendar.google.com"));
        const isPurelyInternal = !allParticipants.some(p => !isInternalEntity(p));
        const hasExternalAttendee = !isPurelyInternal;

        if (externalFilter && isPurelyInternal) continue;

        // --- NEW AE/SOLO LOGIC ---
        const internalParticipantsList = allParticipants.filter(p => isInternalEntity(p));
        const internalAttendeesCount = internalParticipantsList.length;
        const isSoloMeeting = hasExternalAttendee && (internalAttendeesCount === 1);
        
        const internalEmailsString = internalParticipantsList.join(',');

        if (attendanceFilter === "Show Only Accepted/Maybe") {
            let myStatus = "none"; 
            if (organizerEmail === userEmail) {
                myStatus = "organizer";
            } else if (isTrip) {
                // THE NAVAN FIX: Travel bots inject events that sit as "Awaiting". 
                // If it's a recognized trip, bypass the RSVP check entirely so it doesn't get dropped!
                myStatus = "accepted";
            } else if (event.attendees && event.attendees.length > 0) {
                const me = event.attendees.find(a => a.self || a.email.toLowerCase() === userEmail);
                if (me) {
                    myStatus = me.responseStatus;
                } else {
                    myStatus = "accepted"; 
                }
            } else {
                myStatus = "accepted";
            }
            if (!["accepted", "tentative", "organizer"].includes(myStatus)) continue; 
        }
        
        const hasMainTag = descRaw.includes(tagPrefix);
        let originalSeLead = seLeadTagText ? descRaw.includes(seLeadTagText) : false;
        let originalSeCover = seCoverTagText ? descRaw.includes(seCoverTagText) : false;
        let originalInPerson = false;
        let originalTag = "";
        let hasSETagInDescription = false;
        
        if (hasMainTag) {
            // THE FIX: Stop relying on .startsWith(). 
            // Find the exact location of the prefix, no matter what invisible garbage precedes it.
            const prefixIndex = descRaw.indexOf(tagPrefix);
            
            if (prefixIndex !== -1) {
                // Grab everything sitting *after* the prefix
                let remainder = descRaw.substring(prefixIndex + tagPrefix.length);
                
                // Stop at the first newline to ensure we only grab the tag itself
                let rawContent = remainder.split(/\r?\n/)[0].trim();
                
                hasSETagInDescription = rawContent.includes("(SE)");
                
                // Clean up HTML entities and broken tags
                rawContent = rawContent.replace(/(?:p|br|div|span|hr|li|ul|ol|table|td|tr|th|html|body|a)(?:>|&gt;)/gi, '');
                rawContent = rawContent.replace(/&nbsp;|\u200B/gi, '').trim();
                
                if (modifier && rawContent.includes(modifier)) { originalInPerson = true; }
                
                let lookupKey = rawContent.replace(modifier, "").replace(/;/g, "").trim();
                originalTag = tagMap.get(lookupKey) || "";
            }
        }
        
        let suggestedSeLead = originalSeLead;
        let suggestedSeCover = originalSeCover; 
        let suggestedInPerson = originalInPerson;
        let suggestedTag = originalTag;

        if (locationFilterEnabled && event.location) {
            const locationLower = event.location.toLowerCase();
            const allLocationBlockers = [...virtualExclusions, ...locationExclusions];
            const isBlockedLocation = allLocationBlockers.some(keyword => locationLower.includes(keyword));
            
            if (!isBlockedLocation && locationLower.trim() !== "") {
                if (locationRequiresExternal && !hasExternalAttendee) {
                } else { suggestedInPerson = true; }
            }
        }

        let tagFound = false;
        for (const rule of tagWaterfall) {
            let isMatch = false;
            if (rule.checkTitle && _isMatch(titleSearch, rule.rules)) isMatch = true;
            else if (rule.checkDesc && _isMatch(descSearch, rule.rules)) isMatch = true;

            if (isMatch) {
                suggestedTag = rule.tagName; 
                tagFound = true;
                break; 
            }
        }

        if (useAutoSeLead && _isMatch(fullSearch, autoSeLeadRules)) suggestedSeLead = true;
        if (useAutoSeCover && _isMatch(fullSearch, autoSeCoverRules)) suggestedSeCover = true;
        if (suggestedTag && useAutoInPerson && _isMatch(fullSearch, autoInPersonRules)) { suggestedInPerson = true; }

        const eventStartTime = new Date(event.start.dateTime || event.start.date);
        
        const rowData = [
            titleRaw, eventStartTime, organizerEmail, hasExternalAttendee, 
            suggestedSeCover, suggestedSeLead, suggestedInPerson, suggestedTag, 
            '', 
            `=HYPERLINK("${event.htmlLink}", "Open Event")`, 
            '', 
            originalSeCover, originalSeLead, originalInPerson, originalTag,
            durationInMinutes, 
            nightsCount,
            eventColorId,
            isSoloMeeting,
            internalEmailsString
        ];
        
        if (showTaggedMeetings) {
            if (hasMainTag) { fullyTaggedMeetings.push(rowData); } 
            else { noTagMeetings.push(rowData); }
        } else if (filterForMissingSETags) {
            if (hasSETagInDescription) { /* Hide */ } 
            else if (hasMainTag) { aeTaggedMeetings.push(rowData); } 
            else { noTagMeetings.push(rowData); }
        } else {
            if (!hasMainTag) { noTagMeetings.push(rowData); }
        }
    }
    
    const sheet = file.getSheetByName("Untagged Meetings");
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 0) {
          const range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
          range.clearContent();
          range.clearFormat();
          range.setWrap(false);
    }
    
    const headers = ["Title", "Start Time", "Created By", "External Attendees", "SE Cover", "SE Lead", "In Person", "Add Tag to Meeting", "Synced", "Event Link", "Last Refreshed", "Original Cover", "Original SE", "Original IP", "Original Tag", "Duration (min)", "Nights", "Current Color ID", "Solo Meeting", "Internal Emails"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    const finalList = [...noTagMeetings, ...aeTaggedMeetings, ...fullyTaggedMeetings];
    if (finalList.length > 0) { 
        sheet.getRange(2, 1, finalList.length, headers.length).setValues(finalList); 
        const numRows = finalList.length;
        const formulas = [];
        for (let i = 2; i <= numRows + 1; i++) {
              formulas.push([`=AND(E${i}=L${i}, F${i}=M${i}, G${i}=N${i}, H${i}=O${i})`]);
        }
        sheet.getRange(2, 9, numRows, 1).setFormulas(formulas);
    }
    else { sheet.getRange("A2").setValue("No meetings found using your current criteria."); }
    
    _formatUntaggedSheet(sheet, headers, noTagMeetings, aeTaggedMeetings, fullyTaggedMeetings, { filterForMissingSETags, showTaggedMeetings, externalFilter });
    
    sheet.hideColumns(18, 3); 
    updateTagDropdownsAndColors(true); 
    
  } catch (e) { SpreadsheetApp.getUi().alert(e.message); }
}

function updateMeetingTag(e, isSingleEditMode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = isSingleEditMode ? e.range.getSheet() : ss.getSheetByName("Untagged Meetings");
  try {
    if (isSingleEditMode) {
      const range = e.range;
      if (range.getRow() <= 1) return;
      const sheetName = sheet.getName();
      if (sheetName !== "Untagged Meetings") return;
      return; 
    }

    ss.toast("Processing batch... Please wait.");
    const tagPrefix = "Verkada Meeting Tag: ";
    
    const seLeadTag = getConfig("SELeadTagText") || "";
    const seCoverTag = getConfig("SECoverageTagText") || ""; 
    const modifier = getConfig("InPersonModifier") || "";
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { ss.toast("No data to process."); return; }
    
    const tagsSheet = ss.getSheetByName("Tags");
    // UPDATED: Now reading up to Column H (index 7 in 0-based array) to get Color Name
    const tagsRaw = tagsSheet.getRange("B2:H" + tagsSheet.getLastRow()).getValues();
    const validTags = tagsRaw.filter(row => row[2] && row[2].trim() !== "");
    
    const tagMap = new Map(validTags.map(row => [row[1], row[2]])); 
    // NEW: Map Tag Name -> Color Name (Col H is index 6 relative to B2:H)
    const tagColorMap = new Map(validTags.map(row => [row[1], row[6]])); 

    const sheetData = sheet.getRange(2, 1, lastRow - 1, 18).getValues(); 
    
    const userEmail = Session.getActiveUser().getEmail();
    let calendarId;
    try {
        const cal = CalendarApp.getCalendarById(userEmail);
        calendarId = cal ? cal.getId() : CalendarApp.getDefaultCalendar().getId();
    } catch(e) { calendarId = CalendarApp.getDefaultCalendar().getId(); }
    
    const eventTimes = sheetData.map(row => new Date(row[1]).getTime()).filter(t => !isNaN(t));
    if (eventTimes.length === 0) { ss.toast("No valid dates found to process."); return; }
    const minTime = new Date(Math.min(...eventTimes));
    const maxTime = new Date(Math.max(...eventTimes));
    maxTime.setHours(23, 59, 59);

    let allEventsForMap = [];
    let pageToken = null;
    do {
      const response = Calendar.Events.list(calendarId, { timeMin: minTime.toISOString(), timeMax: maxTime.toISOString(), singleEvents: true, maxResults: 2500, pageToken: pageToken });
      if (response.items) allEventsForMap = allEventsForMap.concat(response.items);
      pageToken = response.nextPageToken;
    } while (pageToken);

    const eventMap = new Map();
    allEventsForMap.forEach(event => eventMap.set(`${event.summary}_${new Date(event.start.dateTime || event.start.date).getTime()}`, event));

    const rowsToDelete = [];
    let updatedCount = 0;
    let errorCount = 0; 
    
    for (let i = 0; i < sheetData.length; i++) {
      const rowData = sheetData[i];
      const currentSeCover = rowData[4]; 
      const currentSeLead = rowData[5];
      const currentInPerson = rowData[6];
      const currentTag = rowData[7];
      const originalSeCover = rowData[11];
      const originalSeLead = rowData[12];
      const originalInPerson = rowData[13];
      const originalTag = rowData[14];
      
      // NEW: Retrieve stored color ID from hidden Col R
      const originalColorId = rowData[17] ? String(rowData[17]) : "";

      // Check if text changed OR if color needs update
      let needsColorUpdate = false;
      let targetColorId = originalColorId;

      if (currentTag && currentTag.trim() !== "" && currentTag !== "<Clear Tag>") {
          const colorName = tagColorMap.get(currentTag);
          if (colorName && EVENT_COLOR_MAP.hasOwnProperty(colorName)) {
             targetColorId = EVENT_COLOR_MAP[colorName];
             if (targetColorId !== originalColorId) {
                needsColorUpdate = true;
             }
          }
      }

      const hasChanged = (currentSeCover !== originalSeCover) || (currentSeLead !== originalSeLead) || (currentInPerson !== originalInPerson) || (currentTag !== originalTag) || needsColorUpdate;

      if (hasChanged) {
        const title = rowData[0];
        const startTime = new Date(rowData[1]).getTime();
        const lookupKey = `${title}_${startTime}`;
        const eventToUpdate = eventMap.get(lookupKey);

        if (eventToUpdate) {
          let description = eventToUpdate.description || "";

          if (seLeadTag) description = description.replace(new RegExp(seLeadTag.trim(), 'gi'), '');
          if (seCoverTag) description = description.replace(new RegExp(seCoverTag.trim(), 'gi'), '');
          
          description = description.replace(new RegExp(tagPrefix + ".*", 'gi'), '');

          description = description.trim();
          const isHtml = /<(?!(?:https?|tel):)[a-z][\s\S]*>/i.test(description);
          const breakChar = isHtml ? "<br>" : "\n";

          if (currentSeLead && seLeadTag) description += breakChar + breakChar + seLeadTag;
          if (currentSeCover && seCoverTag) description += breakChar + seCoverTag; 
          
          if (currentTag && currentTag.trim() !== "" && currentTag !== "<Clear Tag>") {
            let abbreviatedTag = tagMap.get(currentTag); 
            if (!abbreviatedTag) continue; 
            
            if (currentInPerson && abbreviatedTag) {
                let parts = abbreviatedTag.split(';').map(s => s.trim()).filter(s => s !== "");
                if (parts.length > 0) {
                    let first = parts[0];
                    let suffix = first.includes('(SE)') ? ' (SE)' : '';
                    let core = first.replace('(SE)', '').trim();
                    parts[0] = `${core}${modifier}${suffix}`;
                    abbreviatedTag = parts.join('; ') + ';';
                }
            }
            description += breakChar + breakChar + tagPrefix + abbreviatedTag;
          }
          
          try {
              // 1. Update Description
              Calendar.Events.patch({ description: description }, calendarId, eventToUpdate.id, { sendUpdates: "none" });
              
              // 2. NEW: Update Color (if required)
              if (needsColorUpdate) {
                  // We must use 'patch' with the specific colorId field
                  Calendar.Events.patch({ colorId: targetColorId }, calendarId, eventToUpdate.id);
              }

              rowsToDelete.push(i + 2);
              updatedCount++;
          } catch (e) {
              console.log(`Permission Error patching event "${title}": ${e.message}`);
              errorCount++;
          }
        }
      }
    }

    if (updatedCount > 0) {
      const numDataRows = lastRow - 1;
      if (updatedCount === numDataRows && rowsToDelete.length === numDataRows) {
        sheet.getRange(2, 1, numDataRows, sheet.getLastColumn()).clearContent();
      } else {
        rowsToDelete.sort((a, b) => b - a).forEach(rowNum => sheet.deleteRow(rowNum));
      }
      
      let msg = `${updatedCount} meeting(s) updated.`;
      if (errorCount > 0) msg += ` (Skipped ${errorCount} locked events).`;
      ss.toast(msg);
      
    } else if (errorCount > 0) {
      ss.toast(`Failed to update ${errorCount} meetings due to permissions.`);
    } else {
      ss.toast("No changes were detected to apply.");
    }

  } catch (err) { SpreadsheetApp.getUi().alert(err.message); }
}

// =================================================================
// #region 3a. REGEX PARSERS (v4.4.0 - Special Character Safety)
// =================================================================

/**
 * PARSER: Converts "vce(!coverage), [demo]" into rule objects.
 * FIX: Now safely handles keywords containing literal () or [] brackets.
 */
function _parseMatchRule(rawString) {
  if (!rawString) return [];
  
  const parts = String(rawString).split(',').map(s => s.trim()).filter(Boolean);
  const rules = [];

  parts.forEach(part => {
    // 1. Check if an exclusion group exists at the very end of the string
    // This regex safely ignores parentheses that are part of the core keyword
    const exclusionMatch = part.match(/\(!([^)]+)\)$/);
    
    let keyword = part;
    let excludes = [];

    if (exclusionMatch) {
        // Strip the exclusion part off the end to leave just the keyword
        keyword = part.replace(exclusionMatch[0], '').trim().toLowerCase();
        excludes = exclusionMatch[1].split('|').map(x => x.trim().toLowerCase()).filter(Boolean);
    } else {
        keyword = keyword.toLowerCase();
    }

    if (keyword) {
        rules.push({ keyword: keyword, excludes: excludes });
    }
  });
  return rules;
}

/**
 * STRICT MATCHER (v2.9.6): Checks text against parsed rules using WORD BOUNDARIES.
 * FIX: Safely escapes special characters and allows colons, hyphens, and slashes.
 */
function _isMatch(text, rules) {
  if (!text || !rules || rules.length === 0) return false;

  for (const rule of rules) {
    const escapedKeyword = rule.keyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    
    // Expanded boundary list includes colons, hyphens, slashes, and ampersands
    const regex = new RegExp(`(^|\\s|[.,;:!?'"()\\[\\]{}<>\\/\\|&+=*~_\\-])${escapedKeyword}(?=\\s|[.,;:!?'"()\\[\\]{}<>\\/\\|&+=*~_\\-]|$)`, 'i');

    if (regex.test(text)) {
        let isExcluded = false;
        if (rule.excludes.length > 0) {
            for (const excl of rule.excludes) {
                if (text.includes(excl)) {
                    isExcluded = true;
                    break;
                }
            }
        }
        if (!isExcluded) return true;
    }
  }
  return false;
}
// #endregion

// =================================================================
// #region 4. HELPER & UTILITY FUNCTIONS (v3.13.0 - Frozen Columns Fix)
// =================================================================

function getConfig(label) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Config");
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === label) return data[i][1];
  }
  return null;
}

function _formatMinutesToHours(totalMinutes) {
  if (totalMinutes === 0) return "0min";
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  let result = "";
  if (hours > 0) result += `${hours}hr`;
  if (minutes > 0) result += ` ${minutes}min`;
  return result.trim();
}

function _getFilterList(listName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheetByName("Filter Lists");
  if (!listSheet) return []; 
  
  // THE FIX: Safety valve to prevent "0 rows" crash if the sheet is empty
  const lastRow = listSheet.getLastRow();
  if (lastRow < 2) return []; 

  const headers = listSheet.getRange('A1:1').getValues()[0];
  const colIndex = headers.findIndex(h => h && h.toString().trim() === listName);
  if (colIndex === -1) return []; 
  
  const columnData = listSheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  return columnData.flat().filter(String).map(value => value.toString().trim());
}

/**
 * Safely renames an existing sheet to "OLD [Name]" to preserve data.
 * If "OLD [Name]" already exists, it is deleted to prevent duplicates.
 */
function _backupOldSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName(sheetName);

  if (currentSheet) {
    const backupName = "OLD " + sheetName;
    const existingBackup = ss.getSheetByName(backupName);

    // 1. If an OLD version already exists, delete it (Nuclear option for the backup)
    if (existingBackup) {
      ss.deleteSheet(existingBackup);
    }

    // 2. Rename the current sheet to the backup name
    currentSheet.setName(backupName);
    
    // 3. Visual formatting to indicate it is an archive
    currentSheet.setTabColor("#cdcd00"); // Dark Gold color
    
    // 4. HIDE SHEET TOGGLE
    // By default, we keep the backup visible so you can copy data easily.
    // To make it HIDDEN automatically, simply remove the "//" from the start of the line below:
    
    // currentSheet.hideSheet();
    
    ss.toast(`Backed up existing '${sheetName}' to '${backupName}'`);
  }
}

function _formatUntaggedSheet(sheet, headers, noTagMeetings, aeTaggedMeetings, fullyTaggedMeetings, configs) {
    // FIX: Freeze first 1 Row (Header) and first 2 Columns (Title, Date)
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2); 
    
    // Header Style
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight("bold")
         .setFontFamily('Poppins')
         .setFontSize(11)
         .setBackground("#283e4d")
         .setFontColor("white")
         .setVerticalAlignment("middle")
         .setHorizontalAlignment("center")
         .setWrap(true); 
    
    const lastRow = sheet.getLastRow();
    const startRow = 2;
    if (lastRow > 1) {
        // Force Black/Normal text on data
        const dataRange = sheet.getRange(startRow, 1, lastRow - 1, headers.length);
        dataRange.setFontFamily('Poppins')
                 .setFontSize(10)
                 .setFontColor("black") 
                 .setFontWeight("normal") 
                 .setVerticalAlignment("middle");

        // Date Column Format
        sheet.getRange(startRow, 2, lastRow - 1, 1)
             .setNumberFormat("yyyy-mm-dd hh:mm AM/PM")
             .setHorizontalAlignment("left");
        
        // External Attendees Color Logic
        const extValues = sheet.getRange(startRow, 4, lastRow - 1, 1).getValues();
        sheet.getRange(startRow, 4, lastRow - 1, 1)
             .setBackgrounds(extValues.map(([val]) => val === true ? ['#d9ead3'] : ['#f4cccc']));
        
        // Insert Checkboxes
        sheet.getRange(startRow, 5, lastRow - 1, 3).insertCheckboxes();
    }
    
    sheet.getRange("K1").setValue(`Last Refreshed: ${new Date().toLocaleString()}`);

    // SET COLUMN WIDTHS
    sheet.setColumnWidth(1, 400); // Title
    sheet.setColumnWidth(2, 150); // Date
    sheet.setColumnWidth(3, 230); // Created By
    sheet.setColumnWidth(4, 150); // External
    sheet.setColumnWidth(5, 80);  // SE Cover
    sheet.setColumnWidth(6, 80);  // SE Lead
    sheet.setColumnWidth(7, 80);  // In Person
    sheet.setColumnWidth(8, 300); // Tag Dropdown
    sheet.setColumnWidth(9, 65);  // Synced
    sheet.setColumnWidth(10, 100); // Link
    sheet.setColumnWidth(11, 350); // Refreshed Tag
    
    // Hide Metadata
    sheet.hideColumns(12, 6); 
    
    // Row Background Colors
    if (configs.showTaggedMeetings) {
        if (noTagMeetings.length > 0) sheet.getRange(startRow, 1, noTagMeetings.length, 3).setBackground('#f4cccc');
    } else if (configs.filterForMissingSETags) {
        if (noTagMeetings.length > 0) sheet.getRange(startRow, 1, noTagMeetings.length, 3).setBackground('#f4cccc');
        if (aeTaggedMeetings.length > 0) sheet.getRange(startRow + noTagMeetings.length, 1, aeTaggedMeetings.length, 3).setBackground('#fff2cc');
    } else {
        if (noTagMeetings.length > 0) sheet.getRange(startRow, 1, noTagMeetings.length, 3).setBackground('#f4cccc');
    }
    
    if (fullyTaggedMeetings.length > 0) {
        const offset = noTagMeetings.length + aeTaggedMeetings.length;
        sheet.getRange(startRow + offset, 1, fullyTaggedMeetings.length, 3).setBackground('#d9ead3');
    }

    // Auto-Tag Suggestion Borders
    const SUGGESTION_COLOR = "#283e4d"; 
    const BORDER_STYLE = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;

    if (lastRow > 1) {
        const checkRange = sheet.getRange(startRow, 5, lastRow - 1, 11).getValues(); 
        for (let i = 0; i < checkRange.length; i++) {
            const row = checkRange[i];
            const visibleRow = i + startRow;
            const suggCover = row[0];
            const suggSE    = row[1];
            const suggIP    = row[2];
            const suggTag   = row[3];
            const origCover = row[7]; 
            const origSE    = row[8]; 
            const origIP    = row[9]; 
            const origTag   = row[10];
            
            if (suggCover !== origCover && suggCover === true) sheet.getRange(visibleRow, 5).setBorder(true, true, true, true, false, false, SUGGESTION_COLOR, BORDER_STYLE);
            if (suggSE !== origSE && suggSE === true) sheet.getRange(visibleRow, 6).setBorder(true, true, true, true, false, false, SUGGESTION_COLOR, BORDER_STYLE);
            if (suggIP !== origIP && suggIP === true) sheet.getRange(visibleRow, 7).setBorder(true, true, true, true, false, false, SUGGESTION_COLOR, BORDER_STYLE);
            if (suggTag !== origTag && suggTag && suggTag !== "") sheet.getRange(visibleRow, 8).setBorder(true, true, true, true, false, false, SUGGESTION_COLOR, BORDER_STYLE);
        }
    }
}

/**
 * Locks or unlocks the "SE Name" Config field based on the Bridge Script's Radar setting.
 */
/**
 * Locks or unlocks the "SE Name" Config field based on the Bridge Script's Radar setting.
 */
function _toggleRadarConfigFields(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  if (!configSheet) return;

  // Find the SE Name row
  const finder = configSheet.createTextFinder("SE Name").matchEntireCell(true).findNext();
  if (!finder) return;

  const targetCell = finder.offset(0, 1);
  
  // Check if Radar is explicitly enabled in the Bridge Script
  const isRadarEnabled = options && options.features && options.features.radar === true;

  if (!isRadarEnabled) {
      // 1. Radar is OFF: Wipe old validation FIRST, set value, then lock it down
      const disabledText = "LEAVE BLANK \n(Radar Disabled)";
      
      targetCell.clearDataValidations(); // Prevents the cell violation error
      targetCell.setValue(disabledText);
      targetCell.setBackground("#f4cccc").setFontColor("#cc0000"); 

      // Create a new dropdown with ONLY the disabled text
      const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList([disabledText], true)
          .setAllowInvalid(false)
          .build();
      targetCell.setDataValidation(rule);
      
  } else {
      // 2. Radar is ON: Clear the warning and restore their team dropdown!
      const currentVal = targetCell.getValue();
      if (String(currentVal).includes("LEAVE BLANK")) {
          targetCell.setValue("");
          targetCell.setBackground("white").setFontColor("black");
      }
      
      // Safely rebuild the original SE Name dropdown using the Bridge Script config
      if (options && options.team && options.team.members && options.team.members.length > 0) {
          const rule = SpreadsheetApp.newDataValidation()
              .requireValueInList(options.team.members, true)
              .setAllowInvalid(false)
              .build();
          targetCell.setDataValidation(rule);
      } else {
          targetCell.clearDataValidations();
      }
  }
}
// #endregion

// =================================================================
// #region 5. DYNAMIC QUARTER & ADMIN FUNCTIONS (v3.4.1)
// =================================================================

function _getDatesForVerkadaFQ(fiscalYear, quarter) {
  let startMonth, endMonth, startYear, endYear;
  const calendarYear = fiscalYear - 1; 
  
  switch (quarter) {
    case 1: 
        startMonth = 1; endMonth = 4; startYear = calendarYear; endYear = calendarYear; break;
    case 2: 
        startMonth = 4; endMonth = 7; startYear = calendarYear; endYear = calendarYear; break;
    case 3: 
        startMonth = 7; endMonth = 10; startYear = calendarYear; endYear = calendarYear; break;
    case 4: 
        startMonth = 10; endMonth = 1; startYear = calendarYear; endYear = calendarYear + 1; break;
    default: throw new Error("Invalid quarter.");
  }
  
  const endDate = new Date(endYear, endMonth, 0); endDate.setHours(23, 59, 59, 999);
  return { start: new Date(startYear, startMonth, 1), end: endDate };
}

function _getVerkadaFiscalQuarter(date) {
  const month = date.getMonth(); 
  const year = date.getFullYear();
  const fy = month >= 1 ? year + 1 : year; 
  const q = (month >= 1 && month <= 3) ? 1 : (month >= 4 && month <= 6) ? 2 : (month >= 7 && month <= 9) ? 3 : 4;
  return { fiscalYear: fy, quarter: q, name: `FY${fy.toString().slice(-2)} Q${q}` };
}

function updateQuarterDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  if (!configSheet) throw new Error("'Config' sheet not found.");
  const finder = configSheet.createTextFinder("QuarterOverride").matchEntireCell(true).findNext();
  if (!finder) { ss.toast("Error: Label 'QuarterOverride' not found."); return; }
  
  const dropdownCell = finder.offset(0, 1);
  let quarters = ["None"];
  let today = new Date();
  for (let i = 0; i < 5; i++) {
    const fq = _getVerkadaFiscalQuarter(today);
    let name = fq.name;
    if (i === 0) name += " (Current)";
    quarters.push(name);
    const dates = _getDatesForVerkadaFQ(fq.fiscalYear, fq.quarter);
    today = new Date(dates.start.getTime() - 1);
  }
  finder.offset(0, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(quarters, true).setAllowInvalid(false).build());
  ss.toast("Quarter dropdown updated.");
}

/**
 * UPDATED (v3.4.1): Now targets Column H (Index 8) for dropdowns.
 */
function updateTagDropdownsAndColors(silentMode = false) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tagsSheet = ss.getSheetByName("Tags");
    const mainSheet = ss.getSheetByName("Untagged Meetings");
    if (!tagsSheet || !mainSheet) { if(!silentMode) ss.toast("Error: Sheets not found."); return; }
    
    if (!silentMode) ss.toast("Updating tag dropdowns and colors...");

    // Color Map (Pastels)
    const COLOR_HEX_MAP = {
      "Lavender": "#d0e0ff", "Sage": "#b6d7a8", "Grape": "#d5a6bd", "Flamingo": "#f4cccc",
      "Banana": "#ffe599", "Tangerine": "#f9cb9c", "Peacock": "#9fc5e8", "Graphite": "#d9d9d9",
      "Blueberry": "#6d9eeb", "Basil": "#93c47d", "Tomato": "#e06666"
    };

    const lastTagRow = tagsSheet.getLastRow();
    // Read B:H (Col 2 to 8) - We need Column B (Section) to detect headers
    const tagRange = tagsSheet.getRange(2, 2, lastTagRow - 1, 8); 
    const tagValues = tagRange.getValues();
    const tagBackgrounds = tagRange.getBackgrounds();

    const tagNames = [["<Clear Tag>"]]; 
    const tagColors = [["#ffffff"]]; 
    const newBackgrounds = [];

    // --- LOOP THROUGH ROWS ---
    for (let i = 0; i < tagValues.length; i++) {
        const sectionName = tagValues[i][0]; // Index 0 = Column B
        const tagName = tagValues[i][1];     // Index 1 = Column C
        
        // 1. GATHER DROPDOWN DATA
        if (tagName && tagName.trim() !== "") {
            tagNames.push([tagName]);
            tagColors.push([tagBackgrounds[i][1]]); 
        }

        // 2. DETERMINE COLUMN I BACKGROUND
        if (sectionName && sectionName !== "") {
            // FIX: It's a Section Header! Force Dark Blue.
            newBackgrounds.push(["#283e4d"]); 
        } else {
            // It's a Data Row. Check for Color Label.
            const colorLabel = tagValues[i][6]; // Index 6 = Column H
            if (colorLabel && COLOR_HEX_MAP[colorLabel]) {
                newBackgrounds.push([COLOR_HEX_MAP[colorLabel]]);
            } else {
                newBackgrounds.push(["#f3f3f3"]); // Default Grey
            }
        }
    }

    // Apply Backgrounds to Tags Sheet (Col I)
    if (newBackgrounds.length > 0) {
        tagsSheet.getRange(2, 9, newBackgrounds.length, 1).setBackgrounds(newBackgrounds);
    }

    // --- RE-APPLY CONDITIONAL FORMATTING (To keep live updates working) ---
    const tagsSheetRules = [];
    const colIRange = tagsSheet.getRange(2, 9, tagsSheet.getMaxRows() - 1, 1);
    
    Object.keys(COLOR_HEX_MAP).forEach(colorName => {
        const hex = COLOR_HEX_MAP[colorName];
        const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`=H2="${colorName}"`)
            .setBackground(hex)
            .setRanges([colIRange])
            .build();
        tagsSheetRules.push(rule);
    });
    
    tagsSheet.setConditionalFormatRules(tagsSheetRules);

    // --- UPDATE UNTAGGED MEETINGS ---
    const dropdownRange = mainSheet.getRange(2, 8, mainSheet.getMaxRows() - 1, 1);
    dropdownRange.clearDataValidations();

const helperSheetName = "_TagListHelper";
    let helperSheet = ss.getSheetByName(helperSheetName);
    
    // 1. Create or clear it (but don't hide it yet!)
    if (!helperSheet) { 
        helperSheet = ss.insertSheet(helperSheetName); 
    } else { 
        helperSheet.clear(); 
    }

    // 2. Write the dropdown data to it
    if (tagNames.length > 0) {
        const helperRange = helperSheet.getRange(1, 1, tagNames.length, 1);
        helperRange.setValues(tagNames);
        helperRange.setBackgrounds(tagColors);
        dropdownRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(helperRange, true).setAllowInvalid(false).build());
    }

    // 3. Force it into hiding AFTER all data operations are done
    helperSheet.hideSheet();

    // 4. Safety Net: Bounce the user's view back to the main sheet 
    // This prevents the UI from glitching and forcing the hidden sheet back open.
    if (mainSheet) {
        mainSheet.activate();
    }

    // Formatting for Untagged Meetings
    mainSheet.clearConditionalFormatRules();
    let rules = [];
    const dropdownRangeForRules = mainSheet.getRange(2, 8, mainSheet.getMaxRows() - 1, 1);
    
    tagNames.forEach((nameArr, i) => {
        const color = tagColors[i][0];
        if (nameArr[0] !== "") {
            rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(nameArr[0]).setBackground(color).setRanges([dropdownRangeForRules]).build());
        }
    });

    const syncedRange = mainSheet.getRange(2, 9, mainSheet.getMaxRows() - 1, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("TRUE").setBackground('#d9ead3').setFontColor('#38761d').setRanges([syncedRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("FALSE").setBackground('#f4cccc').setFontColor('#cc0000').setRanges([syncedRange]).build());
    mainSheet.setConditionalFormatRules(rules);

    if (!silentMode) {
        ss.toast("Tag dropdowns and colors updated.");
        try {
            const statusCell = tagsSheet.getRange("A1");
            statusCell.check();
            statusCell.setBackground("#d9ead3"); 
        } catch (e) { }
    }
}

function validateTagKeywords() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tagsSheet = ss.getSheetByName("Tags");
    if (!tagsSheet) { ss.toast("Error: 'Tags' sheet not found."); return; }
    ss.toast("Validating keywords and labels...");

    const lastRow = tagsSheet.getLastRow();
    if (lastRow < 2) { ss.toast("No data to validate."); return; }

    // Range B to I (8 columns). 
    // B=0, C=1, D=2, E=3 (Keywords), F=4, G=5, H=6, I=7
    const data = tagsSheet.getRange(2, 2, lastRow - 1, 8).getValues(); 

    const keywordTracker = new Map();
    const duplicateKeywords = new Map();
    
    // Maps Color -> First Custom Name found
    const colorDefinitions = new Map(); 
    const colorConflicts = [];

    data.forEach((row) => {
        const tagName = row[1];  // Col C: Tag Name
        const keywords = row[3]; // Col E: Keywords (THIS WAS THE BUG! It was row[2] before)
        const color = row[6];    // Col H: Color Label
        const customName = row[7] ? row[7].toString().trim().toLowerCase() : ""; // Col I: Custom Name

        // 1. KEYWORD CHECK
        // Force keywords to String to prevent crashes if someone types a pure number
        if (tagName && keywords && String(keywords).trim() !== "") {
            String(keywords).split(',').map(k => {
                // Strip out "(!exclusion)" before checking duplicates
                return k.replace(/\(![^)]+\)$/, '').trim().toLowerCase();
            }).filter(String).forEach(k => {
                if (keywordTracker.has(k) && keywordTracker.get(k) !== tagName) {
                    if (!duplicateKeywords.has(k)) {
                        duplicateKeywords.set(k, [keywordTracker.get(k)]);
                    }
                    duplicateKeywords.get(k).push(tagName);
                } else { 
                    keywordTracker.set(k, tagName); 
                }
            });
        }

        // 2. COLOR LABEL CHECK
        if (color && color !== "") {
            const colorKey = color; 
            const meaning = customName !== "" ? customName : color;

            if (colorDefinitions.has(colorKey)) {
                const existingMeaning = colorDefinitions.get(colorKey);
                if (existingMeaning !== meaning) {
                    colorConflicts.push(`- Color "${color}" used for: "${existingMeaning}" AND "${meaning}"`);
                }
            } else {
                colorDefinitions.set(colorKey, meaning);
            }
        }
    });

    let msg = "";
    if (duplicateKeywords.size > 0) {
        msg += "DUPLICATE KEYWORDS:\n";
        for (const [k, t] of duplicateKeywords.entries()) { msg += `- "${k}": ${[...new Set(t)].join(', ')}\n`; }
        msg += "\n";
    }

    if (colorConflicts.length > 0) {
        msg += "COLOR LABEL CONFLICTS (First match wins in Visuals):\n";
        msg += colorConflicts.join("\n");
    }

    if (msg !== "") {
        SpreadsheetApp.getUi().alert("Validation Warnings", msg, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
        ss.toast("Success! All keywords and labels are valid.");
    }
}

function onEditTagsSheet(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() === "Tags") {
    try {
      const statusCell = sheet.getRange("A1");
      const r = e.range.getRow();
      const c = e.range.getColumn();

      // If user clicks A1, don't loop infinitely. Just let them toggle it if they want.
      if (r === 1 && c === 1) return;
      
      // If editing Data (Rows 2+, Cols 2-9)
      if (r > 1 && c >= 2 && c <= 9) { 
          statusCell.uncheck();
          statusCell.setBackground("#f4cccc"); // Red
      }
    } catch (err) { }
  }
}
// #endregion

// =================================================================
// #region 6. DASHBOARDS AND VISUALS
// =================================================================

function generateDashboard(options = {}) {
  // --- FIX START: LOAD MASTER CONFIG ---
  if (!options || Object.keys(options).length === 0) {
      if (typeof APP_CONFIG !== 'undefined') { options = APP_CONFIG; }
  }

  let showDiscrepancyTable = false;
  if (typeof options === 'boolean') {
      showDiscrepancyTable = options;
  } else if (typeof options === 'object') {
      showDiscrepancyTable = options.dashboard?.showDiscrepancyTable === true;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Untagged Meetings");
  const radarSheet = ss.getSheetByName("All SE Radar Meetings CFQ") || ss.getSheetByName("SFDC - Tagged SE Radar Meetings"); 
  const tagsSheet = ss.getSheetByName("Tags");
  const visualsSheetName = "Visuals";
  
  let myEmail = "";
  try { myEmail = Session.getActiveUser().getEmail().toLowerCase(); } catch(e) {}
  
  let visualsSheet = ss.getSheetByName(visualsSheetName);
  if (!visualsSheet) { 
      visualsSheet = ss.insertSheet(visualsSheetName); 
  } else { 
      visualsSheet.getCharts().forEach(c => visualsSheet.removeChart(c));
      visualsSheet.clear(); 
      const maxRows = visualsSheet.getMaxRows();
      visualsSheet.showRows(1, maxRows);
      const fullRange = visualsSheet.getRange(1, 1, maxRows, 1);
      for (let i = 0; i < 8; i++) {
        try { fullRange.shiftRowGroupDepth(-1); } catch(e) { break; }
      }
  }

  // --- EXPAND SHEET SIZE ---
  const currentMaxCols = visualsSheet.getMaxColumns();
  if (currentMaxCols < 50) {
      visualsSheet.insertColumnsAfter(currentMaxCols, 50 - currentMaxCols);
  }

  visualsSheet.showColumns(1, visualsSheet.getMaxColumns());
  visualsSheet.getRange("A1:Z1000").setFontFamily("Poppins").setFontSize(10);
  visualsSheet.setFrozenRows(6); 

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert("No data found in Untagged Meetings sheet."); return; }
  
  // --- 0. PREPARE COLOR MAPPING ---
  const EVENT_COLOR_MAP_REVERSE = {};
  Object.keys(EVENT_COLOR_MAP).forEach(key => {
      EVENT_COLOR_MAP_REVERSE[EVENT_COLOR_MAP[key]] = key;
  });

  const CHART_COLOR_PALETTE = {
      "Lavender": "#7986cb", "Sage": "#33b679", "Grape": "#8e24aa", "Flamingo": "#e67c73",
      "Banana": "#f09300",   "Tangerine": "#f4511e", "Peacock": "#039be5", "Graphite": "#616161",
      "Blueberry": "#3f51b5", "Basil": "#0b8043", "Tomato": "#d50000", "Default": "#a4bdfc"
  };

  const colorIdToNameMap = new Map();
  const colorIdToHexMap = new Map();

  if (tagsSheet) {
      const tLast = tagsSheet.getLastRow();
      if (tLast > 1) {
          const tData = tagsSheet.getRange(2, 2, tLast - 1, 8).getValues();
          tData.forEach(row => {
              const colorName = row[6]; 
              const customName = row[7]; 
              if (colorName && EVENT_COLOR_MAP[colorName]) {
                  const cId = EVENT_COLOR_MAP[colorName];
                  if (!colorIdToNameMap.has(cId)) {
                      colorIdToNameMap.set(cId, customName && customName.toString().trim() !== "" ? customName : colorName);
                      colorIdToHexMap.set(cId, CHART_COLOR_PALETTE[colorName] || "#999999");
                  }
              }
          });
      }
  }

  // --- 1. PROCESS CALENDAR DATA (20 COLUMNS) ---
  const data = dataSheet.getRange(2, 1, lastRow - 1, 20).getValues();
  
  let total = data.length, external = 0, seLead = 0, seCover = 0, inPerson = 0, totalMinutes = 0, totalNights = 0, totalSolo = 0;
  const groupedStats = new Map();
  
  // AE TRACKING SETUP
  const supportedConfig = options.supportedEntities;
  const trackAEs = supportedConfig?.enabled && supportedConfig?.teams;
  const aeData = {}; 
  const aeTeamTotals = {}; 
  const aeGlobalTags = {}; 
  const aeLookup = {};

  if (trackAEs) {
      Object.keys(supportedConfig.teams).forEach(team => {
          aeTeamTotals[team] = { total: 0, time: 0, tags: {}, members: [] };
          supportedConfig.teams[team].forEach(email => {
              const cleanEmail = email.toLowerCase().trim();
              aeLookup[cleanEmail] = team;
              aeData[cleanEmail] = { team: team, total: 0, time: 0, tags: {} };
              aeTeamTotals[team].members.push(cleanEmail);
          });
      });
  }

  const getBaseName = (name) => name.replace(/\s*\(.*?\)\s*$/, "").trim();

  data.forEach(row => {
    // TIE-BREAKER: Grab the organizer's email from Column C (Index 2)
    const organizerEmail = row[2] ? String(row[2]).trim().toLowerCase() : "";
    const isExt = row[3] === true;
    const isCover = row[4] === true;
    const isLead = row[5] === true;
    const isIP = row[6] === true;
    const tagName = row[7];
    const duration = Number(row[15]) || 0; 
    const nights = Number(row[16]) || 0;
    const isSolo = row[18] === true;
    const internalEmailsString = row[19] || ""; 

    if (isExt) external++;
    if (isCover) seCover++;
    if (isLead) seLead++;
    if (isIP) inPerson++;
    if (isSolo) totalSolo++;
    totalMinutes += duration;
    totalNights += nights;
    
    // --- AE TRACKING & AD-HOC COVERAGE LOGIC ---
    if (trackAEs && internalEmailsString !== "" && isExt) {
        const cleanTag = (tagName && tagName.trim() !== "") ? getBaseName(tagName) : "Untagged";
        const emailArray = internalEmailsString.split(',').map(e => e.trim().toLowerCase());
        const matchedAEs = emailArray.filter(email => aeLookup[email]);

        if (matchedAEs.length === 1) {
            const email = matchedAEs[0];
            const team = aeLookup[email];
            
            aeData[email].total++;
            aeData[email].time += duration;
            aeData[email].tags[cleanTag] = (aeData[email].tags[cleanTag] || 0) + 1;
            
            aeTeamTotals[team].total++;
            aeTeamTotals[team].time += duration;
            aeTeamTotals[team].tags[cleanTag] = (aeTeamTotals[team].tags[cleanTag] || 0) + 1;
            
            aeGlobalTags[cleanTag] = (aeGlobalTags[cleanTag] || 0) + 1;

        } else if (matchedAEs.length === 0 && isCover) {
            const otherInternals = emailArray.filter(e => e !== myEmail);
            
            if (otherInternals.length > 0) { // Changed to > 0 to handle multiple Verkada attendees
                // TIE-BREAKER: Give credit to the Organizer. If organizer isn't in the list, default to the first rep.
                let coveredEmail = otherInternals[0]; 
                if (otherInternals.includes(organizerEmail)) {
                    coveredEmail = organizerEmail;
                }

                const virtualTeam = "Ad-Hoc SE Coverage";
                
                if (!aeTeamTotals[virtualTeam]) {
                    aeTeamTotals[virtualTeam] = { total: 0, time: 0, tags: {}, members: [] };
                }
                if (!aeData[coveredEmail]) {
                    aeData[coveredEmail] = { team: virtualTeam, total: 0, time: 0, tags: {} };
                    aeTeamTotals[virtualTeam].members.push(coveredEmail);
                }
                
                aeData[coveredEmail].total++;
                aeData[coveredEmail].time += duration;
                aeData[coveredEmail].tags[cleanTag] = (aeData[coveredEmail].tags[cleanTag] || 0) + 1;
                
                aeTeamTotals[virtualTeam].total++;
                aeTeamTotals[virtualTeam].time += duration;
                aeTeamTotals[virtualTeam].tags[cleanTag] = (aeTeamTotals[virtualTeam].tags[cleanTag] || 0) + 1;
                
                aeGlobalTags[cleanTag] = (aeGlobalTags[cleanTag] || 0) + 1;
            }
        }
    }

    // Tag Aggregation
    if (tagName && tagName.trim() !== "") { 
        const baseName = getBaseName(tagName);
        if (!groupedStats.has(baseName)) {
            groupedStats.set(baseName, { 
                variants: new Map(), 
                radar: { count: 0, ip: 0, lead: 0 },
                agg: { count: 0, time: 0, ip: 0, lead: 0, cover: 0, ext: 0, nights: 0, solo: 0 }
            });
        }
        const group = groupedStats.get(baseName);
        group.agg.count += 1;
        group.agg.time += duration;
        group.agg.nights += nights;
        if (isIP) group.agg.ip += 1;
        if (isLead) group.agg.lead += 1;
        if (isCover) group.agg.cover += 1;
        if (isExt) group.agg.ext += 1;
        if (isSolo) group.agg.solo += 1;
        
        if (!group.variants.has(tagName)) {
            group.variants.set(tagName, { count: 0, time: 0, ip: 0, lead: 0, cover: 0 });
        }
        const variant = group.variants.get(tagName);
        variant.count += 1;
        variant.time += duration;
        if (isIP) variant.ip += 1;
        if (isLead) variant.lead += 1;
        if (isCover) variant.cover += 1;
    }
  });
  
  let internal = total - external;

  // --- 2. PROCESS RADAR DATA ---
  if (radarSheet && radarSheet.getLastRow() > 1) {
      const radarData = radarSheet.getDataRange().getValues();
      const headers = radarData[0];
      const typeIdx = headers.indexOf("Meeting Type");
      const ipIdx = headers.indexOf("Verkada Action: In Person Meeting");
      const leadIdx = headers.indexOf("Verkada Action: SE Lead");
      const isTrue = (val) => String(val).toLowerCase() === "true" || val === 1;

      if (typeIdx > -1) {
          for (let i = 1; i < radarData.length; i++) {
              const rTag = radarData[i][typeIdx];
              if (rTag) {
                  const rBase = getBaseName(rTag.toString().trim());
                  if (groupedStats.has(rBase)) {
                      const grp = groupedStats.get(rBase);
                      grp.radar.count += 1;
                      if (ipIdx > -1 && isTrue(radarData[i][ipIdx])) grp.radar.ip += 1;
                      if (leadIdx > -1 && isTrue(radarData[i][leadIdx])) grp.radar.lead += 1;
                  } 
              }
          }
      }
  }

  // --- 3. DYNAMIC SCORECARDS ---
  const fTotal = `=${total} - SUMIF(O11:O, TRUE, I11:I)`;
  const fExternal = `=${external} - SUMIF(O11:O, TRUE, AA11:AA)`;
  const fInternal = `=${internal} - (SUMIF(O11:O, TRUE, I11:I) - SUMIF(O11:O, TRUE, AA11:AA))`;
  const fInPerson = `=${inPerson} - SUMIF(O11:O, TRUE, J11:J)`;
  const fSeLead = `=${seLead} - SUMIF(O11:O, TRUE, K11:K)`;
  const fSeCover = `=${seCover} - SUMIF(O11:O, TRUE, L11:L)`;
  const fNights = `=${totalNights} - SUMIF(O11:O, TRUE, AB11:AB)`;
  const fSolo = `=${totalSolo} - SUMIF(O11:O, TRUE, AD11:AD)`;

  const bg = "#283e4d";
  const fg = "white";

  const row1Cards = [
    { title: "Total Meetings", val: fTotal, col: 1 },
    { title: "External Meetings", val: fExternal, col: 3 },
    { title: "Internal Meetings", val: fInternal, col: 5 },
    { title: "Total Blocked Mtg Time", val: `=TEXT(SUMIF(O11:O, FALSE, AC11:AC)/1440, "[h]""hr ""mm""min""")`, col: 7 }, 
    { title: "Solo Meetings", val: fSolo, col: 9 } 
  ];
  row1Cards.forEach(card => {
    visualsSheet.getRange(3, card.col, 1, 2).merge().setValue(card.title).setFontWeight("bold").setHorizontalAlignment("center").setBackground(bg).setFontColor(fg).setBorder(true, true, true, true, true, true);
    visualsSheet.getRange(4, card.col, 1, 2).merge().setFormula(card.val).setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  });

  const row2Cards = [
    { title: "In Person Mtgs", val: fInPerson, col: 1 },
    { title: "SE Lead Mtgs", val: fSeLead, col: 3 },
    { title: "SE Coverage Mtgs", val: fSeCover, col: 5 }, 
    { title: "Intentionally Left Blank", val: '="-"', col: 7 }, 
    { title: "Nights Away", val: fNights, col: 9 } 
  ];
  row2Cards.forEach(card => {
    visualsSheet.getRange(5, card.col, 1, 2).merge().setValue(card.title).setFontWeight("bold").setHorizontalAlignment("center").setBackground(bg).setFontColor(fg).setBorder(true, true, true, true, true, true);
    visualsSheet.getRange(6, card.col, 1, 2).merge().setFormula(card.val).setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  });

  const scRanges = [visualsSheet.getRange(4, 1, 1, 10), visualsSheet.getRange(6, 1, 1, 10)];
  let scorecardRules = [
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(100).setFontColor("#38761d").setBold(true).setRanges(scRanges).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(100, 150).setFontColor("#bf9000").setBold(true).setRanges(scRanges).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(150).setFontColor("#cc0000").setBold(true).setRanges(scRanges).build()
  ];

  // --- 4. BUILD TABLE DATA ---
  const sortedGroups = [...groupedStats.entries()].sort((a, b) => b[1].agg.count - a[1].agg.count);
  const displayRows = [];     
  const comparisonRows = [];  
  const hiddenMetricRows = []; 
  const groupsToCreate = []; 
  let currentRowIndex = 11;
  const parentRowsIndices = []; 

  sortedGroups.forEach(([baseName, data]) => {
      const isMulti = data.variants.size > 1;
      displayRows.push([
          baseName + (isMulti ? " (Combined)" : ""), 
          data.agg.count, data.agg.ip, data.agg.lead, data.agg.cover,
          _formatMinutesToHours(data.agg.time)
      ]);
      
      const diffCount = Math.round(data.agg.count - data.radar.count);
      const diffIP = Math.round(data.agg.ip - data.radar.ip);
      const diffLead = Math.round(data.agg.lead - data.radar.lead);
      
      comparisonRows.push([
          baseName, 
          data.agg.count, data.radar.count, diffCount,
          data.agg.ip,    data.radar.ip,    diffIP,
          data.agg.lead,  data.radar.lead,  diffLead
      ]);
      
      hiddenMetricRows.push([data.agg.ext, data.agg.nights, data.agg.time, data.agg.solo]);
      
      parentRowsIndices.push(currentRowIndex); 
      const parentRowIdx = currentRowIndex;
      currentRowIndex++;

      if (isMulti) {
          const variants = [...data.variants.entries()].sort((a, b) => b[1].count - a[1].count);
          variants.forEach(([tagName, vStats]) => {
              displayRows.push([
                  "   ↳ " + tagName,
                  vStats.count, vStats.ip, vStats.lead, vStats.cover,
                  _formatMinutesToHours(vStats.time)
              ]);
              hiddenMetricRows.push([0, 0, 0, 0]);
              currentRowIndex++;
          });
          groupsToCreate.push({ start: parentRowIdx + 1, num: variants.length });
      }
  });

  // --- 5. WRITE MAIN TABLE ---
  const mainHeaderRow = 9;
  const mainStartRow = 11;
  let mainEndRow = mainStartRow + displayRows.length;
  let currentBottomRow = mainEndRow;
  
  visualsSheet.getRange(mainHeaderRow, 8, 1, 8).merge().setValue("From Untagged Meetings List and Calendar")
              .setFontWeight("bold").setFontColor(fg).setBackground("#3d9fd2") 
              .setHorizontalAlignment("center").setFontSize(11).setBorder(true, true, true, true, true, true);

  const headers = ["Tag Name", "Total", "IP", "Lead", "Cov", "Block Time", "%", "Hide"];
  visualsSheet.getRange(mainHeaderRow + 1, 8, 1, 8).setValues([headers])
              .setFontWeight("bold").setBackground(bg).setFontColor(fg).setHorizontalAlignment("center");
  
  if (displayRows.length > 0) {
    visualsSheet.getRange(mainStartRow, 8, displayRows.length, 6).setValues(displayRows);
    
    visualsSheet.getRange(mainStartRow, 27, hiddenMetricRows.length, 4).setValues(hiddenMetricRows).setNumberFormat("0");
    
    visualsSheet.getRange(mainStartRow, 15, visualsSheet.getMaxRows() - mainStartRow, 1).removeCheckboxes();
    parentRowsIndices.forEach(idx => visualsSheet.getRange(idx, 15).insertCheckboxes());
    
    const pctFormulas = [];
    for (let i = mainStartRow; i < mainEndRow; i++) {
        pctFormulas.push([`=IF(O${i}="", "-", IF(O${i}=TRUE, "-", I${i}/SUMIF($O$11:$O,FALSE,$I$11:$I)))`]);
    }
    visualsSheet.getRange(mainStartRow, 14, displayRows.length, 1).setFormulas(pctFormulas).setNumberFormat("0.0%");

    visualsSheet.getRange(mainStartRow, 9, displayRows.length, 4).setNumberFormat("0").setHorizontalAlignment("center");
    visualsSheet.getRange(mainStartRow, 13, displayRows.length, 2).setHorizontalAlignment("right");
    visualsSheet.getRange(mainStartRow, 9, displayRows.length, 1).setBackground("#f3f3f3");
    visualsSheet.getRange(mainStartRow, 13, displayRows.length, 2).setBackground("#f3f3f3");
    parentRowsIndices.forEach(idx => visualsSheet.getRange(idx, 8).setFontWeight("bold"));
    visualsSheet.getRange(mainHeaderRow + 1, 8, displayRows.length + 1, 8).setBorder(true, true, true, true, true, true);
  }

  // --- 6. WRITE SUPPORTED AE TRACKING TABLE ---
  if (trackAEs) {
      const aeHeaderRow = Math.max(39, currentBottomRow + 4);
      const aeStartRow = aeHeaderRow + 2;

      const formatName = (email) => {
          try {
              const namePart = email.split('@')[0];
              return namePart.split('.').map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
          } catch(e) { return email; }
      };

      const mainTableTags = sortedGroups.map(x => x[0]);
      const topAeTags = mainTableTags.filter(t => aeGlobalTags[t]); 
            
      const aeHeaderRenameMap = {
          "Best Practice": "BP",
          "Existing Customer Support": "CX Supp",
          "Trial Setup/Config": "Trials",
          "Project Scoping": "Scoping",
          "Verkada-Sponsored Event": "Event",
          "Verkada Internal Team Discussions": "Internal",
          "Deal Related Discussion": "Deal Sync",
          "Existing Customer Check-In / CBR": "CBR",
          "Pricing / Proposal": "Proposal",
          "Partner Onboarding / Training": "PT Training"
      };

      const earlyTags = topAeTags.slice(0, 4);
      const lateTags = topAeTags.slice(4); 

      const displayEarly = earlyTags.map(t => aeHeaderRenameMap[t] || t);
      const displayLate = lateTags.map(t => aeHeaderRenameMap[t] || t);

      while (displayEarly.length < 4) { displayEarly.push(""); }

      const aeTableHeaders = ["Team / AE", "", "Total Mtgs", ...displayEarly, "Time", "% Time", ...displayLate, "Other"];
      const aeTableData = [];
      const aeBoldRows = []; 
      
      let currentRowTracker = 0;

      Object.keys(aeTeamTotals).forEach(team => {
          const tData = aeTeamTotals[team];

          let teamKnownTags = 0;
          const pct = totalMinutes > 0 ? (tData.time / totalMinutes) : 0;
          
          const teamRow = [team, "", tData.total]; 
          
          for (let i = 0; i < 4; i++) {
              if (i < earlyTags.length) {
                  const c = tData.tags[earlyTags[i]] || 0;
                  teamRow.push(c);
                  teamKnownTags += c;
              } else {
                  teamRow.push("");
              }
          }
          
          teamRow.push(_formatMinutesToHours(tData.time));
          teamRow.push(pct);
          
          lateTags.forEach(tag => {
              const c = tData.tags[tag] || 0;
              teamRow.push(c);
              teamKnownTags += c;
          });
          teamRow.push(tData.total - teamKnownTags);
          
          aeTableData.push(teamRow);
          aeBoldRows.push(currentRowTracker); 
          currentRowTracker++;

          tData.members.forEach(email => {
              const aData = aeData[email];
              
              let aKnown = 0;
              const aPct = totalMinutes > 0 ? (aData.time / totalMinutes) : 0;
              const aRow = ["   ↳ " + formatName(email), "", aData.total];
              
              for (let i = 0; i < 4; i++) {
                  if (i < earlyTags.length) {
                      const c = aData.tags[earlyTags[i]] || 0;
                      aRow.push(c);
                      aKnown += c;
                  } else {
                      aRow.push("");
                  }
              }
              
              aRow.push(_formatMinutesToHours(aData.time));
              aRow.push(aPct);
              
              lateTags.forEach(tag => {
                  const c = aData.tags[tag] || 0;
                  aRow.push(c);
                  aKnown += c;
              });
              
              aRow.push(aData.total - aKnown);
              aeTableData.push(aRow);
              currentRowTracker++;
          });
      });

      if (aeTableData.length > 0) {
          visualsSheet.getRange(aeHeaderRow, 1, 1, aeTableHeaders.length).merge()
                    .setValue("Supported Account Executive Breakdown (Excludes internal & multi-AE syncs)")
                    .setFontWeight("bold").setFontSize(11).setFontColor("white").setBackground("#e69138") 
                    .setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);

          visualsSheet.getRange(aeHeaderRow + 1, 1, 1, aeTableHeaders.length).setValues([aeTableHeaders])
                      .setFontWeight("bold").setBackground("#283e4d").setFontColor("white")
                      .setHorizontalAlignment("center");
                      
          visualsSheet.getRange(aeHeaderRow + 1, 1, 1, 2).mergeAcross();

          visualsSheet.getRange(aeStartRow, 1, aeTableData.length, aeTableHeaders.length).setValues(aeTableData);
          visualsSheet.getRange(aeStartRow, 1, aeTableData.length, 2).mergeAcross();
          
          visualsSheet.getRange(aeStartRow, 3, aeTableData.length, aeTableHeaders.length - 2)
                      .setHorizontalAlignment("center").setNumberFormat("0"); 
          visualsSheet.getRange(aeStartRow, 9, aeTableData.length, 1).setNumberFormat("0.0%"); 
          
          visualsSheet.getRange(aeHeaderRow + 1, 1, aeTableData.length + 1, aeTableHeaders.length).setBorder(true, true, true, true, true, true);
          
          for(let i = 0; i < aeTableData.length; i++) {
              if (aeBoldRows.includes(i)) {
                  visualsSheet.getRange(aeStartRow + i, 1, 1, aeTableHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
              } else if (i % 2 !== 0) {
                  visualsSheet.getRange(aeStartRow + i, 1, 1, aeTableHeaders.length).setBackground("#fafafa");
              }
          }
          currentBottomRow = aeStartRow + aeTableData.length;
      }
  }

  // --- 7. WRITE COMPARISON TABLE ---
  SpreadsheetApp.flush(); 
  let finalFormatRules = [...scorecardRules]; 

  if (showDiscrepancyTable && comparisonRows.length > 0) {
      const compHeaderRow = currentBottomRow + 4; 
      const compStartRow = compHeaderRow + 2;
      const compLen = comparisonRows.length;
      
      visualsSheet.getRange(compHeaderRow, 1, 1, 11).merge().setValue("Discrepancy Check (Calendar vs Radar)")
                .setFontWeight("bold").setFontColor("white").setBackground("#6aa84f")
                .setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
      
      const compHeaders = ["Tag Name", "", "Cal Total", "Radar Total", "Diff", "Cal IP", "Radar IP", "Diff", "Cal Lead", "Radar Lead", "Diff"];
      visualsSheet.getRange(compHeaderRow + 1, 1, 1, 11).setValues([compHeaders])
                  .setFontWeight("bold").setBackground("#efefef").setFontColor("black")
                  .setHorizontalAlignment("center");
                  
      visualsSheet.getRange(compHeaderRow + 1, 1, 1, 2).mergeAcross();

      const adjustedCompRows = comparisonRows.map(row => {
          return [row[0], "", row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9]];
      });

      visualsSheet.getRange(compStartRow, 1, compLen, 11).setValues(adjustedCompRows);
      visualsSheet.getRange(compStartRow, 1, compLen, 2).mergeAcross();
      
      visualsSheet.getRange(compStartRow, 3, compLen, 9).setHorizontalAlignment("center").setNumberFormat("0");
      visualsSheet.getRange(compStartRow, 1, compLen, 1).setFontWeight("bold");
      
      visualsSheet.getRange(compHeaderRow + 1, 1, compLen + 1, 11).setBorder(true, true, true, true, true, true);
      
      const diffRanges = [
          visualsSheet.getRange(compStartRow, 5, compLen, 1),
          visualsSheet.getRange(compStartRow, 8, compLen, 1),
          visualsSheet.getRange(compStartRow, 11, compLen, 1)
      ];
      finalFormatRules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setFontColor("#38761d").setBold(true).setBackground("#d9ead3").setRanges(diffRanges).build());
      finalFormatRules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberNotEqualTo(0).whenCellNotEmpty().setFontColor("#cc0000").setBold(true).setBackground("#f4cccc").setRanges(diffRanges).build());
  }

  visualsSheet.setConditionalFormatRules(finalFormatRules);

  groupsToCreate.forEach(grp => {
      try { visualsSheet.getRange(grp.start, 1, grp.num, 1).shiftRowGroupDepth(1); } catch(e) {}
  });

  // --- 8. CHART GENERATION ---
  visualsSheet.getRange("AF10").setValue("Chart Source");
  visualsSheet.getRange("AF11").setFormula(`=FILTER(H11:I, O11:O=FALSE, LEFT(H11:H, 1) <> " ")`);
  
  if (displayRows.length > 0) {
    const chartRange = visualsSheet.getRange("AF11:AG" + (11 + displayRows.length)); 
    const pieChart = visualsSheet.newChart().setChartType(Charts.ChartType.PIE).addRange(chartRange)
      .setOption('title', 'Meeting Tag Breakdown (Aggregated)')
      .setOption('pieSliceText', 'percentage')
      .setOption('pieSliceTextStyle', { fontName: 'Poppins', color: 'white' }) 
      .setOption('is3D', true)
      .setOption('colors', ['#283e4d', '#3d9fd2', '#757475', '#34545e', '#959ea7', '#546e7a', '#78909c', '#63c0f2', '#1f4e6a', '#de6662', '#c45551', '#e68a87'])
      .setOption('titleTextStyle', { fontName: 'Poppins', fontSize: 20, bold: true })
      .setOption('legend', { position: 'right', textStyle: { fontName: 'Poppins', fontSize: 13 } }) 
      .setOption('chartArea', { left: '5%', top: '10%', width: '70%', height: '80%' }) 
      .setOption('height', 600).setOption('width', 700) 
      .setPosition(9, 1, 0, 0).build(); 
    visualsSheet.insertChart(pieChart);
  }

  // Final Cleanup
  const selectedQuarter = getConfig("QuarterOverride");
  let dateString = `Last ${getConfig("DaysBack")} Days`;
  if (selectedQuarter && selectedQuarter !== "None") { dateString = selectedQuarter; }
  visualsSheet.getRange("A1").setValue(`Dashboard for: ${dateString} (Generated: ${new Date().toLocaleString()})`).setFontWeight("bold").setFontSize(12);

  // --- SAFE COLUMN WIDTHS ---
  [1, 2, 3, 4, 5, 6, 7].forEach(c => visualsSheet.setColumnWidth(c, 100)); 
  visualsSheet.setColumnWidth(8, 250); 
  visualsSheet.setColumnWidth(9, 60);  
  visualsSheet.setColumnWidth(10, 85); 
  visualsSheet.setColumnWidth(11, 60); 
  visualsSheet.setColumnWidth(12, 50); 
  visualsSheet.setColumnWidth(13, 90); 
  visualsSheet.setColumnWidth(14, 60); 
  visualsSheet.setColumnWidth(15, 50); 
  for (let c = 16; c <= 40; c++) { visualsSheet.setColumnWidth(c, 85); } 


  visualsSheet.hideColumns(27, 20); 
  repairSheetStructure(options); 
  ss.setActiveSheet(visualsSheet);
}

// =================================================================
// #region 6a. TEAM CALENDAR CRAWLER & MANAGER DASHBOARD (API V1.8)
// =================================================================

/**
 * 1. DATE RESOLVER & NAME PARSER
 */
function _resolveManagerDashboardDates() {
    const qOverride = getConfig("QuarterOverride");
    let start = new Date(), end = new Date();
    
    if (qOverride && qOverride !== "None" && qOverride !== "") {
        // Normalize the string (uppercase, remove spaces)
        const cleanOverride = String(qOverride).toUpperCase().replace(/\s+/g, '');
        
        // Extract Quarter (Q1, Q2, Q3, Q4)
        const qMatch = cleanOverride.match(/(Q[1-4])/);
        const quarter = qMatch ? qMatch[1] : null;

        // Extract Year (e.g., FY27 -> 2027)
        const yrMatch = cleanOverride.match(/FY(\d{2,4})/);
        let year = new Date().getFullYear();
        const month = new Date().getMonth(); // 0-11
        
        if (yrMatch) {
            // Parse FY27 to 2027, then subtract 1 so it starts in Feb 2026
            let parsedYear = parseInt(yrMatch[1], 10);
            parsedYear = parsedYear < 100 ? 2000 + parsedYear : parsedYear; 
            year = parsedYear - 1; 
        } else {
            // Original Smart Year Offset (only if no explicit FY is provided)
            if (quarter === "Q4" && month < 6) year -= 1;
            if (quarter === "Q3" && month < 3) year -= 1;
        }

        // Apply dates based on the extracted Quarter (Verkada FY: Feb 1 Start)
        if (quarter === "Q1") { 
            start = new Date(year, 1, 1);    // Feb 1
            end = new Date(year, 3, 30);     // Apr 30
        }
        else if (quarter === "Q2") { 
            start = new Date(year, 4, 1);    // May 1
            end = new Date(year, 6, 31);     // Jul 31
        }
        else if (quarter === "Q3") { 
            start = new Date(year, 7, 1);    // Aug 1
            end = new Date(year, 9, 31);     // Oct 31
        }
        else if (quarter === "Q4") { 
            start = new Date(year, 10, 1);   // Nov 1
            end = new Date(year + 1, 0, 31); // Jan 31 (of the NEXT year)
        } else {
             // Fallback just in case they type something completely unreadable
             const dBack = parseInt(getConfig("DaysBack"), 10) || 30;
             const dAhead = parseInt(getConfig("DaysAhead"), 10) || 7;
             start.setDate(start.getDate() - dBack);
             end.setDate(end.getDate() + dAhead);
        }
    } else {
        // Normal rolling window if no override is set
        const dBack = parseInt(getConfig("DaysBack"), 10) || 30;
        const dAhead = parseInt(getConfig("DaysAhead"), 10) || 7;
        start.setDate(start.getDate() - dBack);
        end.setDate(end.getDate() + dAhead);
    }
    
    start.setHours(0,0,0,0);
    end.setHours(23,59,59,999);
    return { start, end };
}

function _parseNameFromEmail(email) {
    if (!email) return "Unknown SE";
    const localPart = String(email).split('@')[0]; 
    const cleanLocal = localPart.replace(/-backup$/i, ''); 
    const parts = cleanLocal.split(/[._]/);
    const capitalized = parts.map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase());
    return capitalized.join(' ');
}

/**
 * 2. CRAWLER ENGINE (Backend DB Sync)
 */
function refreshTeamCalendarDatabase(appConfig) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!appConfig || !appConfig.team || !appConfig.team.enabled) return false;

    const teamEmails = appConfig.team.members || [];
    if (teamEmails.length === 0) return false;

    ss.toast("Crawling team calendars via API...", "Syncing", 15);

    const dbSheetName = "_TeamMeetingsDB";
    let dbSheet = ss.getSheetByName(dbSheetName);
    if (!dbSheet) { dbSheet = ss.insertSheet(dbSheetName).hideSheet(); } 
    else { dbSheet.clear(); }

    const headers = ["SE Name", "SE Email", "Title", "Start Time", "Duration (min)", "Tag", "Is Trip", "Nights Away", "External Attendees", "Attendee List", "Event ID", "Event Link"];
    dbSheet.getRange(1, 1, 1, headers.length).setValues([headers])
           .setFontWeight("bold").setBackground("#283e4d").setFontColor("white");
    dbSheet.setFrozenRows(1);

    const dates = _resolveManagerDashboardDates();
    const tagMap = _buildTagDictionary();
    const nightsAwayRules = _parseMatchRule(_getFilterList("NightsAwayKeywords").join(","));
    const tagPrefix = "Verkada Meeting Tag: ";
    const mod = getConfig("InPersonModifier") || "";
    
    let allParsedEvents = [];
    let seenEventIds = new Set(); 

    for (let i = 0; i < teamEmails.length; i++) {
        const email = teamEmails[i];
        const seName = _parseNameFromEmail(email);
        ss.toast(`Crawling calendar ${i + 1} of ${teamEmails.length}...\n(${seName})`, "Progress", 5);

        try {
            let pageToken = null;
            do {
                const response = Calendar.Events.list(email, {
                    timeMin: dates.start.toISOString(), timeMax: dates.end.toISOString(),
                    singleEvents: true, maxResults: 2500, pageToken: pageToken,
                    fields: "items(id,summary,description,start,end,attendees,organizer,htmlLink),nextPageToken" 
                });

                if (response.items) {
                    for (const event of response.items) {
                        if (seenEventIds.has(event.id)) continue;
                        seenEventIds.add(event.id);

                        const title = event.summary || "";
                        let descRaw = (event.description || "").replace(/<br\s*\/?>/gi, '\n').replace(/<p>/gi, '\n').replace(/<\/p>/gi, '\n').replace(/<(?!(?:https?|tel):)[^>]*>/gi, ''); 
                        const fullSearch = (title + " " + descRaw).toLowerCase();
                        
                        // 1. Tag Scrubbing 
                        let assignedTag = "Untagged";
                        const hasMainTag = descRaw.includes(tagPrefix);
                        
                        if (hasMainTag) {
                            const prefixIndex = descRaw.indexOf(tagPrefix);
                            if (prefixIndex !== -1) {
                                let remainder = descRaw.substring(prefixIndex + tagPrefix.length);
                                let rawContent = remainder.split(/\r?\n/)[0].trim();
                                
                                rawContent = rawContent.replace(/(?:p|br|div|span|hr|li|ul|ol|table|td|tr|th|html|body|a)(?:>|&gt;)/gi, '');
                                rawContent = rawContent.replace(/&nbsp;|\u200B/gi, '').trim();
                                
                                let lookupKey = rawContent.replace(mod, "").replace(/;/g, "").trim();
                                assignedTag = tagMap.get(lookupKey) || "Unknown Tag";
                            }
                        }

                        let isTrip = _isMatch(fullSearch, nightsAwayRules);
                        let nightsAwayCount = 0;
                        let hasExternal = false; 

                        const orgEmail = event.organizer && event.organizer.email ? String(event.organizer.email).toLowerCase().trim() : "";
                        const attendees = (event.attendees || []).map(a => String(a.email).toLowerCase().trim());
                        const allParticipants = [...new Set([orgEmail, ...attendees])].filter(Boolean);
                        
                        const isInternal = em => em.endsWith("@verkada.com") || em.includes("calendar.google.com");
                        hasExternal = allParticipants.some(p => !isInternal(p));

                        // 2. The Guardrail
                        if (isTrip && !hasExternal && allParticipants.length > 2) {
                            isTrip = false;
                        }

                        // 3. SMART NOISE FILTER 
                        if (!hasExternal && !isTrip && assignedTag === "Untagged") continue; 

                        if (isTrip) {
                            if (event.start && event.start.date && event.end && event.end.date) {
                                const sDate = new Date(event.start.date);
                                const eDate = new Date(event.end.date);
                                const diffDays = Math.round((eDate - sDate) / (1000 * 60 * 60 * 24));
                                nightsAwayCount = diffDays > 0 ? diffDays - 1 : 0;
                            } else {
                                nightsAwayCount = 0; 
                            }
                        }

                        const attendeeString = allParticipants.join(", ");
                        
                        // FIX: Safely construct the hyperlink formula by escaping any hidden quotes
                        let eventLink = "No Link Available";
                        if (event.htmlLink) {
                            const safeUrl = String(event.htmlLink).replace(/"/g, '""');
                            eventLink = `=HYPERLINK("${safeUrl}", "Event Link")`;
                        }

                        let durationMin = 0;
                        if (event.start && event.start.dateTime) {
                            const s = new Date(event.start.dateTime).getTime();
                            const e = new Date(event.end.dateTime).getTime();
                            durationMin = Math.round((e - s) / 60000);
                        }

                        allParsedEvents.push([
                            seName, email, title, event.start.dateTime || event.start.date, durationMin, 
                            assignedTag, isTrip, nightsAwayCount, hasExternal, attendeeString, event.id, eventLink
                        ]);
                    }
                }
                pageToken = response.nextPageToken;
            } while (pageToken);
        } catch (err) { console.warn(`Skipped ${email}: ${err.message}`); }
    }

    if (allParsedEvents.length > 0) {
        dbSheet.getRange(2, 1, allParsedEvents.length, headers.length).setValues(allParsedEvents);
        
        // Force formatting AFTER pasting data so Google Sheets cannot auto-detect the wrong format!
        dbSheet.getRange("D:D").setNumberFormat("yyyy-mm-dd hh:mm AM/PM"); 
        dbSheet.getRange("E:E").setNumberFormat("0"); 
    }
    return true;
}


/**
 * 3. DASHBOARD GENERATOR (Decoupled Frontend UI V1.8)
 */
function generateManagerDashboard(appConfig) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!appConfig || !appConfig.team || !appConfig.team.enabled) return;

    const dbSheetName = "_TeamMeetingsDB";
    const dbSheet = ss.getSheetByName(dbSheetName);
    
    if (!dbSheet || dbSheet.getLastRow() < 2) { 
        SpreadsheetApp.getUi().alert("No team data found! Please run the 'TEST_SyncBackendData' tool first.");
        return; 
    }

    const dashSheetName = "Manager Team Analysis";
    let sheet = ss.getSheetByName(dashSheetName);
    if (!sheet) { sheet = ss.insertSheet(dashSheetName); } 
    
    let existingReps = new Map();
    const oldData = sheet.getDataRange().getValues();
    let headerRowIndex = -1;
    for(let r=0; r<oldData.length; r++) { if(oldData[r][0] === "SE Name") { headerRowIndex = r; break; } }
    if (headerRowIndex !== -1) {
        const oldHeaders = oldData[headerRowIndex];
        const repsColIdx = oldHeaders.indexOf("Reps");
        if (repsColIdx !== -1) {
            for(let r=headerRowIndex+1; r<oldData.length; r++) {
                const sName = oldData[r][0];
                const sReps = oldData[r][repsColIdx];
                if (sName && sReps !== "") existingReps.set(sName, sReps);
            }
        }
    }

    sheet.clear();
    const charts = sheet.getCharts();
    for (let i = 0; i < charts.length; i++) { sheet.removeChart(charts[i]); }
    sheet.setHiddenGridlines(true); 

    const dbData = dbSheet.getDataRange().getValues();

    const headerRenameMap = {
        "Best Practice": "BP",
        "Existing Customer Support": "CX Supp",
        "Trial Setup/Config": "Trials",
        "Project Scoping": "Scoping",
        "Verkada-Sponsored Event": "Event", 
        "Existing Customer Check-In / CBR": "CBR",
        "Pricing / Proposal": "Proposal",
        "Partner Onboarding / Training": "Partner Training", 
        "Verkada Internal Team Discussions": "Internal",
        "Deal Related Discussion": "Deal Sync",
        "Dedicated Discovery Call": "Discovery",
        "RFP / RFI / Security Questionnaire": "RFP"
    };

    const seStats = new Map(); 
    const teamWideTagCounts = new Map();
    let totalTeamMeetings = 0;
    let totalTeamNightsAway = 0;

    for (let i = 1; i < dbData.length; i++) {
        const row = dbData[i];
        const seName = row[0]; 
        const durationMin = Number(row[4]) || 0; 
        const rawTag = row[5]; 
        
        const isTrip = row[6] === true || String(row[6]).toUpperCase() === "TRUE"; 
        const nightsAwayCount = Number(row[7]) || 0; 
        const hasExternal = row[8] === true || String(row[8]).toUpperCase() === "TRUE"; 
        
        if (!seStats.has(seName)) { seStats.set(seName, { total: 0, untagged: 0, time: 0, tags: {} }); }
        const stats = seStats.get(seName);

        // 1. Process Trips first
        if (isTrip) {
            totalTeamNightsAway += nightsAwayCount;
            continue; 
        }

        // --- TAG PRE-PROCESSING ---
        let cleanTag = rawTag.endsWith(';') ? rawTag.slice(0, -1).trim() : rawTag;
        let baseTag = cleanTag.replace(/\s*\([^)]+\)$/, '').trim();

        // --- 2. THE BLACKLIST ---
        // Add tags here that should NEVER be counted, regardless of external attendees
        const blacklistedTags = ["Other", "Personal"]; 
        if (blacklistedTags.includes(baseTag)) {
            continue;
        }

        // --- 3. THE INTERNAL WHITELIST ---
        // Add tags here that SHOULD be counted, even if they are internal-only
        const internalExceptions = [
            "AE Enablement" 
        ];
        let isInternalException = internalExceptions.includes(baseTag); 

        // --- 4. THE BOUNCER ---
        if (!hasExternal && !isInternalException) {
            continue; 
        }

        // --- 5. COUNT METRICS ---
        stats.total++;
        stats.time += durationMin; 
        totalTeamMeetings++;

        if (rawTag === "Untagged" || rawTag === "Unknown Tag") {
            stats.untagged++;
            continue;
        }

        let displayTag = headerRenameMap[baseTag] || baseTag;
        stats.tags[displayTag] = (stats.tags[displayTag] || 0) + 1;
        teamWideTagCounts.set(displayTag, (teamWideTagCounts.get(displayTag) || 0) + 1);
    }

    // Determine Top Performers
    let topSE = "N/A"; let topSECount = 0;
    let leastUntagged = Infinity; let leastUntaggedSE = "N/A";
    let maxWalks = 0; let siteWalkerSE = "N/A";
    let maxTime = 0; let mostTimeSE = "N/A"; 
    
    // --- NEW SCORECARD VARIABLES ---
    let maxAeSifu = 0; let aeSifuSE = "N/A"; 
    let maxPartnerSifu = 0; let partnerSifuSE = "N/A";
    let maxCustomerSuperstar = 0; let customerSuperstarSE = "N/A";
    
    const sortedSEs = Array.from(seStats.keys()).sort();

    for (const name of sortedSEs) {
        const stats = seStats.get(name);
        
        // Standard Metrics
        if (stats.total > topSECount) { topSECount = stats.total; topSE = name; }
        if (stats.time > maxTime) { maxTime = stats.time; mostTimeSE = name; } 
        if (stats.total > 0 && stats.untagged < leastUntagged) { leastUntagged = stats.untagged; leastUntaggedSE = name; }
        
        // Tag-Specific Metrics
        const walks = stats.tags["Site Walk"] || 0;
        if (walks > maxWalks) { maxWalks = walks; siteWalkerSE = name; }
        
        const aeEnablementCount = stats.tags["AE Enablement"] || 0;
        if (aeEnablementCount > maxAeSifu) { maxAeSifu = aeEnablementCount; aeSifuSE = name; }
        
        // FIX: Must use the names defined in the 'headerRenameMap' to find the count!
        const partnerCount = stats.tags["Partner Training"] || 0;
        if (partnerCount > maxPartnerSifu) { maxPartnerSifu = partnerCount; partnerSifuSE = name; }
        
        const customerCount = (stats.tags["CX Supp"] || 0) + (stats.tags["CBR"] || 0);
        if (customerCount > maxCustomerSuperstar) { maxCustomerSuperstar = customerCount; customerSuperstarSE = name; }
    }
    
    if (leastUntagged === Infinity) leastUntagged = 0; 
    if (maxAeSifu === 0) aeSifuSE = "None"; 
    if (maxPartnerSifu === 0) partnerSifuSE = "None";
    if (maxCustomerSuperstar === 0) customerSuperstarSE = "None";

    // Format the max time cleanly
    let formattedTime = "0min";
    if (maxTime > 0) {
        const hrs = Math.floor(maxTime / 60);
        const mins = maxTime % 60;
        formattedTime = hrs > 0 ? `${hrs}hr ${mins}min` : `${mins}min`;
    }

    const startRow = 32; 
    const rStart = startRow + 2; 
    const rEnd = startRow + 1 + sortedSEs.length; 
    const dynamicAvgFormula = `=IF(MAX(C${rStart}:C${rEnd})>0, INDEX(A${rStart}:A${rEnd}, MATCH(MAX(C${rStart}:C${rEnd}), C${rStart}:C${rEnd}, 0)) & " (" & TEXT(MAX(C${rStart}:C${rEnd}), "0.1") & ")", "Needs Rep Data")`;

    // THE REORDERED SCORECARDS
    const scorecards = [
        // Top Row (Rows 2 & 3)
        { title: "Total Team Mtgs", value: totalTeamMeetings, color: "#000000" },               
        { title: "Top SE (Most Mtgs)", value: `${topSE} (${topSECount})`, color: "#38761d" },   
        { title: "Highest Avg/Rep", value: dynamicAvgFormula, color: "#6aa84f", isFormula: true }, 
        { title: "Least Untagged", value: `${leastUntaggedSE} (${leastUntagged})`, color: "#990000" }, 
        
        // Second Row (Rows 5 & 6)
        { title: "Total Nights Away", value: totalTeamNightsAway, color: "#b45f06" },           
        { title: "Most Time in Mtgs", value: `${mostTimeSE} (${formattedTime})`, color: "#1155cc" }, 
        { title: "Site Walker", value: `${siteWalkerSE} (${maxWalks})`, color: "#38761d" },     
        { title: "AE Sifu", value: `${aeSifuSE} (${maxAeSifu})`, color: "#674ea7" },
        
        // Third Row (Rows 8 & 9)
        { title: "Partner Sifu", value: `${partnerSifuSE} (${maxPartnerSifu})`, color: "#e69138" }, // Orange
        { title: "Customer Superstar", value: `${customerSuperstarSE} (${maxCustomerSuperstar})`, color: "#d50000" } // Red
    ];
    
    // Dynamically pad the grid so the row always finishes perfectly flush (multiples of 4)
    while (scorecards.length % 4 !== 0) {
        scorecards.push({ title: "Metric Placeholder", value: "BLANK", color: "#cccccc" });
    }

    let cardRow = 2; let cardCol = 10; let count = 0;
    for (let i = 0; i < scorecards.length; i++) {
        const card = scorecards[i];
        sheet.getRange(cardRow, cardCol, 1, 2).merge().setValue(card.title)
             .setFontWeight("bold").setBackground("#f3f3f3").setHorizontalAlignment("center")
             .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
        const valRange = sheet.getRange(cardRow + 1, cardCol, 1, 2).merge()
             .setFontSize(11).setFontWeight("bold").setFontColor(card.color).setHorizontalAlignment("center")
             .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
        if (card.isFormula) { valRange.setFormula(card.value); } else { valRange.setValue(card.value); }
        
        cardCol += 2; count++;
        if (count >= 4) { cardCol = 10; cardRow += 3; count = 0; } 
    }

    let chartDataRow = 1;
    for (const [tag, count] of teamWideTagCounts.entries()) {
        sheet.getRange(`AA${chartDataRow}`).setValue(tag);
        sheet.getRange(`AB${chartDataRow}`).setValue(count);
        chartDataRow++;
    }
    sheet.hideColumn(sheet.getRange("AA1"));
    sheet.hideColumn(sheet.getRange("AB1"));

    if (teamWideTagCounts.size > 0) {
        const customPalette = ['#263a45', '#398bbf', '#737373', '#93a4b0', '#c25a5a', '#546b7a', '#749bb8', '#b55a5a', '#d99c9c', '#b3cbe0', '#e3b2b2'];
        const chartDataRange = sheet.getRange(`AA1:AB${chartDataRow - 1}`);
        const pieChart = sheet.newChart()
            .asPieChart().addRange(chartDataRange).setPosition(2, 1, 0, 0) 
            .setOption('title', 'Team Aggregate: Meeting Types').setOption('is3D', true)
            .setOption('pieSliceText', 'percentage').setOption('pieSliceTextStyle', {color: 'white', fontSize: 13})
            .setOption('colors', customPalette).setOption('legend', {position: 'right', textStyle: {fontSize: 12}})
            .setOption('width', 750).setOption('height', 520)
            .build();
        sheet.insertChart(pieChart);
    }

    const sortedTags = Array.from(teamWideTagCounts.keys()).sort();
    const headers = ["SE Name", "Reps", "Avg/Rep", "Total Mtgs", "Untagged", ...sortedTags];
    let tableData = [];
    for (let idx = 0; idx < sortedSEs.length; idx++) {
        const se = sortedSEs[idx];
        const stats = seStats.get(se);
        const dataRowNumber = startRow + 2 + idx; 
        const repVal = existingReps.has(se) ? existingReps.get(se) : "";
        const avgFormula = `=IF(ISNUMBER(B${dataRowNumber}), D${dataRowNumber}/B${dataRowNumber}, 0)`;
        let row = [se, repVal, avgFormula, stats.total, stats.untagged];
        for (const tag of sortedTags) { row.push(stats.tags[tag] || 0); }
        tableData.push(row);
    }

    if (tableData.length > 0) {
        // --- RICH TEXT HEADER LOGIC ---
        const headerTitle = "UNOFFICIAL Manager Team Meeting Distribution Breakdown (from Calendar)";
        
        const titleRedStyle = SpreadsheetApp.newTextStyle()
            .setForegroundColor("#ff0000") // Bright Red
            .setFontSize(11)           
            .setFontFamily("Poppins")
            .setBold(true)
            .build();
            
        const titleWhiteStyle = SpreadsheetApp.newTextStyle()
            .setForegroundColor("#ffffff") // White
            .setFontSize(11)
            .setFontFamily("Poppins")
            .setBold(true)
            .build();
            
        const richTitle = SpreadsheetApp.newRichTextValue()
            .setText(headerTitle)
            .setTextStyle(0, 10, titleRedStyle) 
            .setTextStyle(10, headerTitle.length, titleWhiteStyle)
            .build();

        sheet.getRange(startRow, 1, 1, headers.length).merge().setRichTextValue(richTitle)
             .setBackground("#1a2b3c").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
        // ----------------------------------

        sheet.getRange(startRow + 1, 1, 1, headers.length).setValues([headers])
             .setFontWeight("bold").setBackground("#283e4d").setFontColor("white").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
        const dataRange = sheet.getRange(startRow + 2, 1, tableData.length, headers.length);
        dataRange.setValues(tableData);
        dataRange.setHorizontalAlignment("center").setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
        sheet.getRange(startRow + 2, 1, tableData.length, 1).setHorizontalAlignment("left").setFontWeight("bold");

        dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

        sheet.getRange(startRow + 2, 2, tableData.length, 1).setBackground("#f8f9fa"); 
        sheet.getRange(startRow + 2, 3, tableData.length, 1).setNumberFormat("0.0"); 
        sheet.getRange(startRow + 2, 4, tableData.length, 1).setNumberFormat("0"); 
        sheet.getRange(startRow + 2, 5, tableData.length, 1).setBackground("#f4cccc").setFontColor("#990000").setFontWeight("bold");

        if (headers.length > 5) {
            const tagDataRange = sheet.getRange(startRow + 2, 6, tableData.length, headers.length - 5);
            const firstCell = `F${startRow + 2}`;
            const colTop1 = `=AND(${firstCell}>0, ${firstCell}=LARGE(F$${rStart}:F$${rEnd}, 1))`;
            const colTop4 = `=AND(${firstCell}>0, ${firstCell}>=LARGE(F$${rStart}:F$${rEnd}, MIN(4, COUNTIF(F$${rStart}:F$${rEnd}, ">0"))), ${firstCell}<LARGE(F$${rStart}:F$${rEnd}, 1))`;
            const ruleGreen = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(colTop1).setBackground("#d9ead3").setRanges([tagDataRange]).build();
            const ruleYellow = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(colTop4).setBackground("#fff2cc").setRanges([tagDataRange]).build();
            sheet.setConditionalFormatRules([ruleGreen, ruleYellow]);
        }

        sheet.setColumnWidth(1, 160); 
        for (let c = 2; c <= headers.length; c++) {
            let dynamicWidth = Math.round((String(headers[c - 1]).length * 7) + 25);
            if (dynamicWidth < 65) dynamicWidth = 65;
            sheet.setColumnWidth(c, dynamicWidth);
        }
    }
    sheet.getRange("A1").activate();
}

/**
 * Helper: Build Tag Dictionary
 */
function _buildTagDictionary() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tagsSheet = ss.getSheetByName("Tags");
    if (!tagsSheet) return new Map();
    const tagsData = tagsSheet.getRange("B2:G" + tagsSheet.getLastRow()).getValues();
    const tagMap = new Map(); 
    tagsData.forEach(row => {
        const tagName = row[1];
        const abbrTag = row[2] ? String(row[2]).replace(/;/g, "").trim() : "";
        if (abbrTag) tagMap.set(abbrTag, tagName);
    });
    return tagMap;
}
// #endregion

// =================================================================
// #region 7. INITIALIZATION & SCAFFOLDING (Include Radar Exclusion Option Flags)
// =================================================================

function initializeAllSheets(options = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableRadar = options.features?.radar !== false; 
  const enableVisuals = options.features?.visuals !== false;

  // 1. Repair Standard Sheets
  repairInstructionsSheet(); 
  repairUntaggedSheet(options); 
  repairConfigSheet(options);    
  repairFilterSheet(options);    
  repairTagsSheet(options); 

  _toggleRadarConfigFields(options);     
  
  // 2. Conditional Radar Logic
  if (enableRadar) {
    repairRadarMainSheet(); 
    repairRadarImportSheet(); 
    repairRadarFilteredSheet(); 
  } else {
    const radarTabs = ["SE Radar Meetings", "All SE Radar Meetings CFQ", "SFDC - Tagged SE Radar Meetings"];
    radarTabs.forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (sheet) sheet.hideSheet();
    });
  }

  // 3. Conditional Visuals Logic (FIXED)
  if (enableVisuals) {
     let visualsSheet = ss.getSheetByName("Visuals");
     if (!visualsSheet) { 
         // If missing, create it!
         visualsSheet = ss.insertSheet("Visuals"); 
     } else {
         // If exists, wipe it clean
         visualsSheet.clear();
     }
  }

  // 4. Sort Sheets
  repairSheetStructure(options); 

  // 5. Final View
  const instructionSheet = ss.getSheetByName("Instructions");
  if (instructionSheet) instructionSheet.activate();

  ss.toast("Full system initialization complete.");
}

function repairSheetStructure(options = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const enableRadar = options.features?.radar !== false; 
  
  // 1. Remember where the user currently is
  const originalSheet = ss.getActiveSheet();

  // 2. Define the Ideal Order
  let sheets = ["Instructions", "Untagged Meetings", "Config", "Filter Lists", "Tags", "Visuals", "Team Visuals"];
  
  if (enableRadar) {
    sheets.push("SE Radar Meetings");
    sheets.push("All SE Radar Meetings CFQ");
  }
  
  // 3. Sort ONLY what exists
  sheets.forEach((name, index) => {
    const sheet = ss.getSheetByName(name);
    if (sheet) { 
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
      sheet.showSheet();
    }
  });

  // 4. Restore original tab focus (if it still exists)
  if (originalSheet) {
      try {
          originalSheet.activate();
      } catch (e) {}
  }
}

// --- NEW FUNCTION ---
function repairInstructionsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Instructions");
  if (!sheet) { sheet = ss.insertSheet("Instructions", 0); }

  sheet.clear();
  sheet.setHiddenGridlines(true); // Cleaner look

  // --- RICH TEXT HEADER LOGIC ---
  const headerRange = sheet.getRange("A1:B1").merge();
  const fullText = "UNOFFICIAL SE Untagged Meetings (UtM) Tool - User Guide";
  
  // 1. Define Style for "UNOFFICIAL" (Bold Red)
  const redStyle = SpreadsheetApp.newTextStyle()
    .setForegroundColor("#ff0000") // Red
    .setBold(true)
    .setFontFamily("Poppins")
    .setFontSize(14)
    .build();

  // 2. Define Style for the rest (Bold White)
  const whiteStyle = SpreadsheetApp.newTextStyle()
    .setForegroundColor("#ffffff") // White
    .setBold(true)
    .setFontFamily("Poppins")
    .setFontSize(14)
    .build();

  // 3. Build the Rich Text Value
  const richText = SpreadsheetApp.newRichTextValue()
    .setText(fullText)
    .setTextStyle(0, 10, redStyle)                 // Indices 0-10: "UNOFFICIAL"
    .setTextStyle(10, fullText.length, whiteStyle) // The rest
    .build();

  // 4. Apply to Cell
  headerRange.setRichTextValue(richText)
             .setBackground("#283e4d") // Dark Blue Background
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle");
             
  sheet.setFrozenRows(1);

  // --- INSTRUCTIONS CONTENT ---
  const instructions = [
    ["Creating an Untagged Meeting List:"],
    ["- Configure your desired settings and filters on the 'Config' tab."],
    ["- You can Update/Refresh your list via the 'Meeting Tools > Save Config and Refresh Untagged List' Menu Dropdown."],
    ["- Add keywords to the 'Filter Lists' and 'Tags' sheets for auto-tag suggestions. You will need to Refresh for those terms to be included"],
    ["- Include Label Colors on each Tag to leverage Google Calendar Labels (Event Color Coding)"],
    ["------ This will be overwritten by 'Clockwise Ideal Day' Colors if enabled on your Calendar"],
    [" "],
    ["- You may also leverage 'Exclusion' logic on a search term to exclude keyword combinations from an auto-tag suggestion"],
    ["------ Example: vce(!coverage) would suggest auto-tags for anything matching 'vce' UNLESS 'coverage' was also found in the title/description"],
    ["------ The exclusions will work in 'Tags' as well as 'Filter Lists' "],
    ["TIP 1: Keep 'MinutesInterval' at 0 (never auto-refresh) if you are planning to bulk edit calendar events"],
    ["TIP 2: An automatic refresh (minutesInterval) will REFRESH your list and you will LOSE un-applied edits on the 'Untagged Meetings' page"],
    [" "],
    ["Submitting Calendar Tag Changes to Your Calendar"],
    ["- When you have made your edits to the 'Untagged Meetings' sheet:"],
    ["------ Select 'Meeting Tools > Apply Changes to Calendar' to update your calendar events in bulk."],
    ["- Updates to Event Tags (via autoTagged suggestions or manually) will show 'Synced: FALSE' if that change has not yet been saved to the event"],
    ["TIP 1: App Scripts limit runtimes to 6 minutes - if timeouts occur due to meeting amount(s) see TIP 2"],
    ["TIP 2: Reduce the amount of DaysBack, and/or set 'SE Tag Mode' & 'ShowTaggedMeetings' to FALSE to reduce query results"],
    [" "],
    ["Generating a Visuals Chart and Meeting Counts"],
    ["- The 'Visuals' sheet is created based on what appears within a populated 'Untagged Meetings' List"],
    ["- Create the Page using the 'Meeting Tools > Generate Dashboard' Menu Dropdown"],
    ["- The creation of the Visuals sheet is meant to be manually run once the 'Untagged Meetings' sheet has been populated"]
  ];

  const startRow = 3;
  // Write text
  sheet.getRange(startRow, 1, instructions.length, 1).setValues(instructions)
        .setFontFamily("Poppins").setFontSize(10).setWrap(true);

  // Styling Headers within the text
  const sectionHeaderIndices = [0, 13, 20]; 
  sectionHeaderIndices.forEach(idx => {
      sheet.getRange(startRow + idx, 1).setFontWeight("bold").setFontColor("#1155cc").setFontSize(11); 
  });

  // Column Width
  sheet.setColumnWidth(1, 1000);
  
  ss.toast("Instructions sheet repaired.");
}


function repairUntaggedSheet(options = {}) { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Untagged Meetings");
  if (!sheet) {
      sheet = ss.insertSheet("Untagged Meetings");
  }

  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  const fullRange = sheet.getRange(1, 1, maxRows, maxCols);

  fullRange.clearContent();
  fullRange.clearFormat();
  fullRange.clearDataValidations(); 
  fullRange.clearNote();
  fullRange.removeCheckboxes();

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2); 

  // Reset columns (Standard 100)
  for (let i = 1; i <= maxCols; i++) {
      sheet.setColumnWidth(i, 100);
  }

  repairSheetStructure(options); 
  ss.toast("Untagged Meetings sheet fully reset.");
}

function repairConfigSheet(options = {}) { // Accept options
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let configSheet = ss.getSheetByName("Config");
  if (!configSheet) configSheet = ss.insertSheet("Config");
  
  configSheet.clear();
  
  // Base Configuration Data
  let configData = [];

  // [OPTIONAL] SFDC SE NAME DROPDOWN BLOCK
  configData.push({ 
      key: "SE Name", 
      val: "", 
      desc: "Select your name to see SFDC meetings on 'SE Radar' Tabs (Leave Blank if no 'SE Radar' Tabs are present)", 
      type: "se_name_dropdown" 
  });

  const standardConfig = [
    { key: "MinutesInterval", val: 30, desc: "How frequently (1, 5, 10, 15,or 30 minutes max) the Untagged Meetings tab auto-refreshes. 0 = No auto-refresh. \nChanges made to config/lists require another 'Save' for suggestions to update.", type: "number" },
    { key: "QuarterOverride", val: "None", desc: "When selected, start/end dates align with Verkada's FY dates. API restricts edits to about 2 quarters in the past\n**Selecting 'None' allows you to to use 'DaysBack' and 'DaysAhead' logic", type: "quarter_dropdown" },
    { key: "DaysBack", val: 30, desc: "How far back from today (in days) to look for untagged meetings (ignored if QuarterOverride selected)", type: "number" },
    { key: "DaysAhead", val: 7, desc: "How far ahead from today (in days) to look for untagged meetings (ignored if QuarterOverride selected)", type: "number" },
    
    { key: "IgnorePhrases", val: true, desc: `CHECKED: Meetings are ignored if they CONTAIN words/phrases located in the IgnorePhrases Filter List.\nUNCHECKED: Ignore this Filter List`, type: "checkbox" },
    { key: "IgnoreExactTitles", val: true, desc: `CHECKED: Exclude Event Titles found in the IgnoreExactTitles Filter List from results\nUNCHECKED: Ignore this Filter List`, type: "checkbox" },
    { key: "IgnoreFromEmails", val: true, desc: `CHECKED: Ignore meetings from specific email addresses located in the IgnoreFromEmails Filter List\nUNCHECKED: Ignore this Filter List`, type: "checkbox" },
    { key: "IgnoreToEmails", val: true, desc: `CHECKED: Ignore meetings sent to specific email addresses located in the IgnoreToEmails Filter List\nUNCHECKED: Ignore this Filter List`, type: "checkbox" },
    
    { key: "IncludeAllDayEvents", val: true, desc: "CHECKED : The script shows all-day events on the \"Untagged Meetings\" list (use with 'NightsAwayKeywords')\nUNCHECKED: The script ignores any all-day events.", type: "checkbox" },
    { key: "NightsAwayKeywords", val: true, desc: `Used with 'IncludeAllDayEvents' filter for ACTUAL Travel Events. Manage keyword suggestions from the NightsAway Filter List`, type: "checkbox" },
    
    { key: "LocationSetsInPerson", val: true, desc: "CHECKED: Suggests an 'In Person' tag on any meeting containing a physical address/location\nUNCHECKED: Does NOT suggest an 'In Person' tag on meetings containing a location", type: "checkbox" },
    { key: "LocationRequiresExternal", val: true, desc: "CHECKED: Location-based 'In Person' suggestions ONLY occur if an External Attendee is present (Filters out internal conference rooms).\nUNCHECKED: Any physical location triggers an 'In Person' suggestion.", type: "checkbox" },
    
    // --- NEW ITEM START --- v2.9.6
    { key: "LocationExclusions", val: true, desc: "CHECKED: If a meeting location contains a word from the LocationExclusions Filter List (e.g. 'Room 404'), the 'In Person' tag will NOT be suggested \nUNCHECKED: Filter List is ignored", type: "checkbox" },
    // --- NEW ITEM END ---

    { key: "AutoInPersonKeywords", val: true, desc: `Scan the full title + description against your AutoInPersonKeywords Filter list.\nIf it finds a match, it will check the 'In Person' checkbox as a suggested calendar event tag MODIFIER`, type: "checkbox" },
    { key: "AutoSELeadKeywords", val: true, desc: `Scan the full title + description against your AutoSELeadKeywords Filter list.\nIf it finds a match, it will check the 'SE Lead' checkbox as a suggested calendar event tag MODIFIER`, type: "checkbox" },
    { key: "AutoSECoverageKeywords", val: true, desc: `Scan the full title + description against your AutoSECoverageKeywords Filter list.\nIf it finds a match, it will check the 'SE Cover' checkbox as a suggested calendar event tag MODIFIER`, type: "checkbox" }, 
    
    { key: "FilterForMissingSETags", val: false, desc: "CHECKED: Show only meetings with no tag, and meetings that have an AE/SFDC implemented tags.\nUNCHECKED: HIDES meetings with any tag (AE/SFDC OR SE Tags)\n**Results will populate with YELLOW if the meeting is missing a ('SE') Tag, and RED if the meeting has NO Tags at all.", type: "checkbox" },
    { key: "ShowTaggedMeetings", val: true, desc: "CHECKED: Displays all tagged meetings (with GREEN) allowing for editing past tagged events\noverrides 'FilterForMissingSETags' behavior to hide \"(SE)\" tagged meetings", type: "checkbox" },
    { key: "ExternalAttendees", val: true, desc: "CHECKED: SHOWS only Meetings with External Attendees\nUNCHECKED: SHOWS Meetings with and without External Attendees (Internal Meetings)", type: "checkbox" },
    
    { key: "InPersonModifier", val: "-P", desc: "This modifier is added to a meeting type if you are In Person instead of joining Virtually.\nThis should not need to be changed but has been put here incase logic changes with Meeting Tags in the future", type: "text" },
    { key: "SELeadTagText", val: "Verkada SE Lead;", desc: "This denotes the text added to a meeting type if you lead the call.\nThis should not need to be changed but has been put here incase logic changes with Meeting Tags in the future", type: "text" },
    { key: "SECoverageTagText", val: "Verkada SE Coverage;", desc: "This denotes the text added to description when 'SE Cover' is checked. \nThis should not need to be changed but has been put here incase logic changes with Meeting Tags in the future", type: "text" },
    { key: "AttendanceFilter", val: "Show Only Accepted/Maybe", desc: "\"Show Only Accepted/Maybe\" = filter by attendance status: \"accepted\", \"tentative\" (Maybe), \"organizer\"\n\"Show All Meetings\" = All meetings are shown -- including declined meetings.", type: "attendance_dropdown" }
  ];

  configData = configData.concat(standardConfig);

  configSheet.getRange("A1:C1").setValues([["Parameter / Filter", "Value", "Description"]])
              .setFontWeight("bold").setBackground("#283e4d").setFontColor("white").setHorizontalAlignment("center");
  configSheet.setFrozenRows(1);

  const outputRows = configData.map(r => [r.key, r.val, r.desc]);
  configSheet.getRange(2, 1, outputRows.length, 3).setValues(outputRows);

  configSheet.getRange("A:C").setFontFamily("Poppins");
  configSheet.getRange("A:A").setFontWeight("bold");
  configSheet.getRange("B:B").setFontWeight("bold").setWrap(true).setHorizontalAlignment("center").setVerticalAlignment("middle");
  configSheet.getRange("C:C").setWrap(true).setVerticalAlignment("middle");
  configSheet.getRange("D:D").setBackground("#283e4d");
  
  configSheet.setColumnWidth(1, 200);
  configSheet.setColumnWidth(2, 150);
  configSheet.setColumnWidth(3, 1200);
  configSheet.setColumnWidth(4, 30);  

  for (let i = 0; i < configData.length; i++) {
    const row = i + 2;
    const cell = configSheet.getRange(row, 2);
    const descCell = configSheet.getRange(row, 3);
    const type = configData[i].type;

    cell.setBackground("#a0d0e9"); 
    descCell.setFontColor("#cc0000");

    if (type === "checkbox") {
      cell.insertCheckboxes();
    } 
    else if (type === "quarter_dropdown") {
      let quarters = ["None"];
      let today = new Date();
      for (let q = 0; q < 5; q++) {
        const fq = _getVerkadaFiscalQuarter(today);
        let name = fq.name;
        if (q === 0) name += " (Current)";
        quarters.push(name);
        const dates = _getDatesForVerkadaFQ(fq.fiscalYear, fq.quarter);
        today = new Date(dates.start.getTime() - 1);
      }
      cell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(quarters, true).build());
    } 
    else if (type === "attendance_dropdown") {
      const options = ["Show Only Accepted/Maybe", "Show All Meetings"];
      cell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(options, true).build());
    } 
    else if (type === "se_name_dropdown") {
      const sfdcSheet = ss.getSheetByName("SFDC - Tagged SE Radar Meetings") || ss.getSheetByName("SFDC - Tagged SE Radar Meeting");
      if (sfdcSheet) {
          const lastRow = sfdcSheet.getLastRow();
          if (lastRow >= 3) {
              try {
                  const fullDataRange = sfdcSheet.getRange(3, 1, lastRow - 2, sfdcSheet.getLastColumn());
                  fullDataRange.sort({column: 1, ascending: true});
              } catch (e) { }
              const range = sfdcSheet.getRange("A3:A");
              const rule = SpreadsheetApp.newDataValidation()
                  .requireValueInRange(range)
                  .build();
              cell.setDataValidation(rule);
          }
      }
    }
    else {
      cell.clearDataValidations();
    }
  }
  // ADDED RADAR SE NAME OPTION TOGGLE (v.2.9.12)
  if (typeof _toggleRadarConfigFields === "function") {
      _toggleRadarConfigFields(options);
  }
  // 2. Restore Order with options passed in:
  repairSheetStructure(options);

  ss.toast("Config sheet repaired.");
}

function repairFilterSheet(options = {}) { // Accept options
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _backupOldSheet("Filter Lists"); 

  let filterSheet = ss.getSheetByName("Filter Lists");
  if (!filterSheet) filterSheet = ss.insertSheet("Filter Lists");
  
  filterSheet.clear();

  // --- DEFINED ORDER ---
  // We explicitly define the order here so Headers and Data ALWAYS match.
  const definedOrder = [
    "IgnorePhrases",
    "IgnoreExactTitles",
    "IgnoreFromEmails",
    "IgnoreToEmails",
    "NightsAwayKeywords",
    "AutoInPersonKeywords",
    "AutoSELeadKeywords",
    "AutoSECoverageKeywords",
    "LocationExclusions" // v2.9.6
  ];

  // 2. Set Headers Directly
  filterSheet.getRange(1, 1, 1, definedOrder.length).setValues([definedOrder]);
  filterSheet.getRange("1:1").setFontWeight("bold").setBackground("#283e4d").setFontColor("white")
              .setFontFamily("Poppins").setVerticalAlignment("middle");
  filterSheet.setFrozenRows(1);

  // 3. Define Defaults
  const filterDefaults = {
    "IgnorePhrases":        [""], 
    "IgnoreExactTitles":    [""], 
    "IgnoreFromEmails":     [""], 
    "IgnoreToEmails":       [""], 
    "NightsAwayKeywords":   ["hotel", "trip", "stay"],  
    "AutoInPersonKeywords": ["sitewalk", "site walk", "vce(!coverage)"], 
    "AutoSELeadKeywords":   [""],
    "AutoSECoverageKeywords": [""],
    "LocationExclusions":   ["room", "huddle", "phone booth", "conf", "internal"] // v2.9.6
  };

  // 4. Map Defaults to the Defined Order
  const defaultColumns = definedOrder.map(key => filterDefaults[key]);

  // --- SMART RESTORE LOGIC ---
  const backupSheet = ss.getSheetByName("OLD Filter Lists");
  if (backupSheet) {
    try {
      const oldHeaders = backupSheet.getRange(1, 1, 1, backupSheet.getLastColumn()).getValues()[0];
      const headerMap = new Map();
      oldHeaders.forEach((h, i) => headerMap.set(h.toString().trim(), i + 1)); 

      definedOrder.forEach((key, idx) => {
         if (headerMap.has(key)) {
            const colIdx = headerMap.get(key);
            const lastRow = backupSheet.getLastRow();
            if (lastRow > 1) {
                const oldData = backupSheet.getRange(2, colIdx, lastRow - 1, 1).getValues().flat();
                defaultColumns[idx] = oldData;
            }
         }
      });
      ss.toast("Restored custom filters from backup.");
    } catch (e) {
      console.log("Filter Restore Error: " + e.message);
    }
  }

  // 5. Write Data
  const maxRows = Math.max(...defaultColumns.map(col => col.length));
  
  const finalFilterData = [];
  for (let r = 0; r < maxRows; r++) {
    const rowData = [];
    for (let c = 0; c < defaultColumns.length; c++) {
      rowData.push(defaultColumns[c][r] || ""); 
    }
    finalFilterData.push(rowData);
  }

  if (finalFilterData.length > 0) {
    filterSheet.getRange(2, 1, finalFilterData.length, finalFilterData[0].length).setValues(finalFilterData);
  }
  
  repairSheetStructure(options);

  // --- 6. RESIZE COLUMNS WITH BUFFER ---
  // First, auto-resize to fit the text
  filterSheet.autoResizeColumns(1, definedOrder.length);

  // Then, add the 7px buffer to every column
  for (let c = 1; c <= definedOrder.length; c++) {
      const currentWidth = filterSheet.getColumnWidth(c);
      filterSheet.setColumnWidth(c, currentWidth + 25); 
  }

  ss.toast("Filter Lists sheet repaired (Columns resized).");
}

function repairTagsSheet(options = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  _backupOldSheet("Tags"); 

  let tagsSheet = ss.getSheetByName("Tags");
  if (!tagsSheet) tagsSheet = ss.insertSheet("Tags");
  
  tagsSheet.clear(); 

  // --- HEADER SCAFFOLDING ---
  // A1 is now Status, I1 is Custom Label
  const tagHeaders = [["Tag Section", "Tag and Color (For Dropdown)", "Abbreviated Tag", "Keywords (auto-tag suggestions)", "Titles", "Descriptions", "Label Color", "Label Name"]];
  tagsSheet.getRange("B1:I1").setValues(tagHeaders)
             .setFontWeight("bold").setBackground("#283e4d").setFontColor("white")
             .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  
  // --- LAYOUT DATA ---
  const tagsLayout = [
    ["Sales Cycle", "", "", "", "", "", "", ""], 
    ["", "Dedicated Discovery Call (AE)", "DS;", "", "", "", "", ""],
    ["", "Demo (AE)", "DM;", "", "", "", "", ""],
    ["", "Best Practice (AE)", "BP;", "", "", "", "", ""],
    ["", "Project Scoping (AE)", "PS;", "", "", "", "", ""],
    ["", "Site Walk (AE)", "SW;", "", "", "", "", ""],
    ["", "Pricing / Proposal (AE)", "PN;", "", "", "", "", ""],
    ["", "Executive Alignment / Board Meeting (AE)", "EX;", "", "", "", "", ""],
    
    ["Customer Success", "", "", "", "", "", "", ""], 
    ["", "Existing Customer Check-In / CBR (AE)", "OC;", "", "", "", "", ""],
    ["", "Existing Customer Support (AE)", "SP;", "", "", "", "", ""],

    ["Solutions Engineering", "", "", "", "", "", "", ""], 
    ["", "Demo (SE)", "DM (SE);", "demo", true, "", "", ""],
    ["", "Best Practice (SE)", "BP (SE);", "best practice", true, "", "", ""],
    ["", "Trial Setup/Config (SE)", "TC (SE);", "trial, setup, config, configuration", true, "", "", ""],
    ["", "Site Walk (SE)", "SW (SE);", "sitewalk, site walk", true, "", "", ""],
    ["", "Floorplans (SE)", "FP (SE);", "floorplan", true, "", "", ""],
    ["", "VCE Training (SE)", "VT (SE);", "vce(!coverage)", true, "", "", ""],
    ["", "Partner Onboarding / Training (SE)", "PT (SE);", "", "", "", "", ""],
    ["", "Industry Conference / Trade Show (SE)", "TS (SE);", "", "", "", "", ""],
    ["", "Verkada-Sponsored Event (SE)", "VM (SE);", "", "", "", "", ""],
    ["", "Existing Customer Support (SE)", "SP (SE);", "", "", "", "", ""],
    ["", "Existing Customer Check-In / CBR (SE)", "OC (SE);", "", "", "", "", ""],
    ["", "AE Enablement (SE)", "AE (SE);", "", "", "", "", ""],
    ["", "RFP / RFI / Security Questionnaire (SE)", "RQ (SE);", "", "", "", "", ""],
    ["", "EBC (SE)", "EB (SE);", "", "", "", "", ""],

    ["Partner", "", "", "", "", "", "", ""], 
    ["", "Account Mapping (PARTNER)", "AM;", "", "", "", "", ""],
    ["", "Deal Related Discussion (PARTNER)", "DR;", "", "", "", "", ""],
    ["", "Partner Onboarding / Training (PARTNER)", "PT;", "", "", "", "", ""],
    ["", "Channel Co-Marketing (PARTNER)", "CM;", "", "", "", "", ""],
    ["", "Partner Demo Certification (PARTNER)", "DC;", "", "", "", "", ""],

    ["Other", "", "", "", "", "", "", ""], 
    ["", "Recruiting (OTHER)", "RT;", "", "", "", "", ""],
    ["", "VCE Training (OTHER)", "VC;", "", "", "", "", ""],
    ["", "Other (OTHER)", "OT;", "", "", "", "", ""],

    ["Personal", "", "", "", "", "", "", ""], 
    ["", "Viper Team", "VPR (SE);", "Viper", true, "", "", ""],
    ["", "Verkada Internal Team Discussions", "INT (SE);", "", "", "", "", ""],
    ["", "Travel Time", "TRVL (SE);", "Flight:, Layover, Drive to, Commute", true, "", "", ""],
    ["", "Overnight", "SLEEP (SE);", "Stay:, Stay At", true, "", "", ""]
  ];

  const dataRange = tagsSheet.getRange(2, 2, tagsLayout.length, 8);
  dataRange.setValues(tagsLayout);
  
  dataRange.setFontColor("black").setBackground("white");

  // --- COLUMN WIDTHS & FORMATTING ---
  tagsSheet.getRange("A:Z").setFontFamily("Poppins").setVerticalAlignment("middle");
  tagsSheet.setFrozenRows(1);
  
  tagsSheet.setColumnWidth(1, 25); // <--- STATUS (Checkbox size)
  tagsSheet.setColumnWidth(2, 150); 
  tagsSheet.setColumnWidth(3, 300); 
  tagsSheet.setColumnWidth(4, 150); // Abbreviated Tag Text
  tagsSheet.setColumnWidth(5, 300); 
  tagsSheet.setColumnWidth(6, 60); 
  tagsSheet.setColumnWidth(7, 100); 
  tagsSheet.setColumnWidth(8, 90); // Label Color
  tagsSheet.setColumnWidth(9, 110); // Label Name Width

  // --- DYNAMIC FORMATTING LOGIC ---
  const sectionPalette = [
    "#efefef", "#ead1dc", "#d9ead3", "#cfe2f3", "#fce5cd", "#d9d2e9", "#fff2cc", "#f4cccc"
  ];
  
  const colorRule = SpreadsheetApp.newDataValidation()
                                  .requireValueInList(Object.keys(EVENT_COLOR_MAP), true)
                                  .setAllowInvalid(true).build();

  const tagColorOverrides = [
    { keyword: "Enterprise", color: "#ff00ff" },
    { keyword: "SE Coverage", color: "#ff00ff" },
    { keyword: "Viper", color: "#ff00ff" },
    { keyword: "Internal", color: "#ff00ff" },
    { keyword: "Travel", color: "#00ffff" },
    { keyword: "Overnight", color: "#ff9900" }
  ];

  let currentSectionColor = sectionPalette[0];
  let paletteIndex = 0;
  
  for (let i = 0; i < tagsLayout.length; i++) {
      const rowNum = i + 2; 
      const sectionName = tagsLayout[i][0]; 
      const tagName = tagsLayout[i][1];       
      const defaultTitle = tagsLayout[i][4]; 
      const defaultDesc = tagsLayout[i][5];  

      if (sectionName !== "") {
          currentSectionColor = sectionPalette[paletteIndex % sectionPalette.length];
          paletteIndex++;
          tagsSheet.getRange(rowNum, 2, 1, 8).setBackground("#283e4d").setFontColor("white").setFontWeight("bold");
      } else if (tagName !== "") {
          // Grey out checkboxes area
          tagsSheet.getRange(rowNum, 5, 1, 3).setBackground("#f3f3f3");
          const checkRange = tagsSheet.getRange(rowNum, 6, 1, 2);
          checkRange.insertCheckboxes();
          
          // Color Dropdown
          tagsSheet.getRange(rowNum, 8).setDataValidation(colorRule).setBackground("#f3f3f3");
          // Custom Label Field (Light Grey default)
          tagsSheet.getRange(rowNum, 9).setBackground("#f3f3f3");
          
          if (defaultTitle === true) tagsSheet.getRange(rowNum, 6).check();
          if (defaultDesc === true) tagsSheet.getRange(rowNum, 7).check();

          let color = currentSectionColor;
          for (const override of tagColorOverrides) {
             if (tagName.includes(override.keyword)) {
               color = override.color;
               break; 
             }
          }
          tagsSheet.getRange(rowNum, 3).setBackground(color);
      }
  }

  // --- NEW STATUS INDICATOR (A1) ---
  const statusCell = tagsSheet.getRange("A1");
  statusCell.insertCheckboxes();
  statusCell.check(); // Default to "Saved"
  statusCell.setBackground("#d9ead3"); // Green
  statusCell.setNote("CHECKED = Saved\nUNCHECKED = Unsaved Changes");

  const tagsFooterStart = tagsLayout.length + 4; 
  const tagsFooterText = [
      ["1. SAVE CHANGES: 'Meeting Tools > Update Tag Dropdowns and Colors'"],
      ["2. CHECKBOX IN A1 INDICATES SAVED STATUS (Green = Saved)"],
      ["3. LABEL NAMES (Column I) ARE NOT SYNCED to your calendar, ONLY LABEL COLORS"],
      ["4. CHANGES MADE HERE WILL NEED YOU TO SAVE THE LIST OR YOU WILL GET 'DATA VALIDATION' ERRORS"]
  ];

  tagsSheet.getRange(tagsFooterStart, 2, tagsFooterText.length, 1).setValues(tagsFooterText);
  tagsSheet.getRange(tagsFooterStart, 2, tagsFooterText.length, 1)
           .setFontColor("#cc0000").setFontWeight("bold");
  
  updateTagDropdownsAndColors(true);
  repairSheetStructure(options); 
  
  ss.toast("Tags sheet repaired.");
}

// =================================================================
//      R A D A R   T A B   C O D E   ( H I D D E N   D A T A )
// =================================================================
// This section handles the "SFDC - Tagged SE Radar Meetings" tab.
// It sets up the ImportRange formulas and formatting.
// =================================================================

function repairRadarImportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "SFDC - Tagged SE Radar Meetings"; // The Hidden Tab
  let sheet = ss.getSheetByName(sheetName);
  
  // 1. Create or Reset Sheet
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clear(); // Wipe clean to ensure formulas don't conflict

  // 2. Define The Formulas (Columns A - O)
  const sourceUrl = "https://docs.google.com/spreadsheets/d/1d2znC5-A7CG787Fsg0AJBwck-nhg2l29tSRjWyGr4Fc/edit?gid=1375013567#gid=1375013567";
  const sourceTab = "'[SFDC] SE Compass - All SE Meetings'!";
  
  const formulas = [
    [`=IMPORTRANGE("${sourceUrl}", "${sourceTab}A:A")`, 
     `=IMPORTRANGE("${sourceUrl}", "${sourceTab}B:B")`, 
     `=IMPORTRANGE("${sourceUrl}", "${sourceTab}C:C")`, 
     `=IMPORTRANGE("${sourceUrl}", "${sourceTab}D:D")`, 
     `=IMPORTRANGE("${sourceUrl}", "${sourceTab}E:E")`, 
     `=IMPORTRANGE("${sourceUrl}", "${sourceTab}F:F")`, 
     `=IMPORTRANGE("${sourceUrl}", "${sourceTab}G:G")`, 
     `=IMPORTRANGE("${sourceUrl}", "${sourceTab}H:H")`, 
     `=IMPORTRANGE("${sourceUrl}", "${sourceTab}I:I")`, 
     `=IMPORTRANGE("${sourceUrl}", "${sourceTab}J:J")`, 
     `=IMPORTRANGE("${sourceUrl}", "${sourceTab}K:K")`, 
     `=IMPORTRANGE("${sourceUrl}", "${sourceTab}L:L")`, 
     `=IMPORTRANGE("${sourceUrl}", "${sourceTab}M:M")`, 
     `=IMPORTRANGE("${sourceUrl}", "${sourceTab}N:N")`, 
     `=IMPORTRANGE("${sourceUrl}", "${sourceTab}O:O")`  
    ]
  ];

  // 3. Write Formulas to Row 3 (Data Start)
  sheet.getRange(3, 1, 1, 15).setFormulas(formulas);

  // 4. Create Headers in Row 2
  // UPDATED: Using the exact list you provided
  const headers = [
    [
      "Full Name", 
      "Subject", 
      "Verkada Action Link Name", 
      "Director: Full Name", 
      "Verkada Action: Actual Meeting Duration", 
      "Verkada Action: Actual Meeting Attendees", 
      "Verkada Reps/AAEs/MDRs Involved", 
      "Verkada Action: Date and Time", 
      "Verkada Action: Participants", 
      "Contacts And Leads Involved", 
      "Verkada Action: Verkada Action", 
      "Meeting Created Date", 
      "Meeting Type", 
      "Verkada Action: In Person Meeting", 
      "Verkada Action: SE Lead"
    ]
  ];
  
  sheet.getRange(2, 1, 1, 15).setValues(headers)
       .setFontWeight("bold")
       .setBackground("#283e4d") // Dark Blue
       .setFontColor("white")
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle");

  // 5. Formatting & View Settings
  sheet.setFrozenRows(2); 
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 300); 
  
  // Hide the sheet as requested
  sheet.hideSheet();
  
  ss.toast("Radar Import sheet (Hidden) has been repaired.");
}

// =================================================================
//      V I S I B L E   R A D A R   T A B   ( C F Q   V I E W )
// =================================================================
// This section manages the "All SE Radar Meetings CFQ" tab.
// It filters the hidden data based on the Config tab (Current User).
// =================================================================

function repairRadarFilteredSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hiddenSheetName = "SFDC - Tagged SE Radar Meetings";
  const visibleSheetName = "All SE Radar Meetings CFQ"; 
  
  let hiddenSheet = ss.getSheetByName(hiddenSheetName);
  let visibleSheet = ss.getSheetByName(visibleSheetName);
  
  // --- CONFIGURATION ---
  // Formula filters the hidden data (starting Row 3) based on Config!B2
  const filterFormula = `=FILTER('${hiddenSheetName}'!A3:O, '${hiddenSheetName}'!A3:A=Config!B2)`;
  // ---------------------

  // 1. Create or Reset Visible Sheet
  if (!visibleSheet) {
    visibleSheet = ss.insertSheet(visibleSheetName);
  }
  visibleSheet.clear(); 

  // 2. DYNAMIC HEADERS (Pull from Hidden Sheet)
  // We grab Row 2 from the hidden sheet (where we established headers live)
  // and paste them into Row 1 of this visible sheet.
  if (hiddenSheet) {
      const sourceHeaders = hiddenSheet.getRange(2, 1, 1, 15).getValues();
      visibleSheet.getRange(1, 1, 1, 15).setValues(sourceHeaders);
  } else {
      // Fallback if hidden sheet is missing (shouldn't happen if initialized in order)
      visibleSheet.getRange("A1").setValue("Error: Source Sheet Missing");
  }

  // 3. Apply Styling (Matches Untagged Meetings)
  const headerRange = visibleSheet.getRange(1, 1, 1, 15);
  headerRange.setFontWeight("bold")
             .setBackground("#283e4d") // Dark Blue
             .setFontColor("white")
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle");

  // 4. Set the Filter Formula in A2
  visibleSheet.getRange("A2").setFormula(filterFormula);

  // 5. Final View Settings
  visibleSheet.setFrozenRows(1);
  visibleSheet.setColumnWidth(1, 150); // Full Name
  visibleSheet.setColumnWidth(2, 300); // Subject width
  
  // Ensure this sheet is Visible
  visibleSheet.showSheet();

  ss.toast("Visible Radar Tab (CFQ) has been repaired.");
}
// =================================================================
//      E N D   V I S I B L E   R A D A R   T A B
// =================================================================

// =================================================================
//      S E   R A D A R   M E E T I N G S   ( M A I N   D A S H )
// =================================================================
// This creates the main interaction tab with the Pivot Table
// and the "Untagged Meetings" list on the right.
// =================================================================

function repairRadarMainSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "SE Radar Meetings";
  const sourceSheetName = "All SE Radar Meetings CFQ";
  
  let sheet = ss.getSheetByName(sheetName);
  const sourceSheet = ss.getSheetByName(sourceSheetName);

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("Error: Source sheet 'All SE Radar Meetings CFQ' is missing. Run the repair for that sheet first.");
    return;
  }

  // --- 1. CRITICAL: PREPARE SOURCE DATA ---
  SpreadsheetApp.flush(); 
  
  const lastRow = Math.max(sourceSheet.getLastRow(), 2); 
  const sourceRange = sourceSheet.getRange(1, 1, lastRow, 15);

  // --- 2. RESET DASHBOARD SHEET ---
  if (!sheet) { sheet = ss.insertSheet(sheetName); }
  sheet.clear(); 
  
  const pivots = sheet.getPivotTables();
  pivots.forEach(p => p.remove());

  // --- 3. SUMMARY SECTION ---
  sheet.getRange("B2:D2").merge()
       .setValue("SE Radar Official Meeting Count (Current Fiscal Quarter)")
       .setFontWeight("bold")
       .setFontFamily("Poppins")
       .setFontSize(11)
       .setBackground("#283e4d")
       .setFontColor("white")
       .setHorizontalAlignment("center")
       .setBorder(true, true, true, true, true, true);

  // Row 6: Total Meetings
  const greyBg = "#efefef";
  sheet.getRange("B6").setValue("Total meetings").setBackground(greyBg).setFontWeight("bold");
  sheet.getRange("C6").setFormula(`=COUNTA('${sourceSheetName}'!C2:C)`).setBackground(greyBg).setFontWeight("bold"); 
  sheet.getRange("D6").setFormula("=C6/$C$6").setBackground(greyBg).setNumberFormat("0%").setFontWeight("bold");

  // Row 5: Untagged Meetings
  const yellowBg = "#fff2cc";
  sheet.getRange("B5").setValue("Untagged meetings").setBackground(yellowBg);
  sheet.getRange("C5").setFormula(`=COUNTIFS('${sourceSheetName}'!C2:C, "<>", '${sourceSheetName}'!M2:M, "")`).setBackground(yellowBg);
  sheet.getRange("D5").setFormula("=C5/$C$6").setBackground(yellowBg).setNumberFormat("0%");

  // Row 4: Tagged Meetings
  const greenBg = "#d9ead3";
  sheet.getRange("B4").setValue("Tagged meetings").setBackground(greenBg);
  sheet.getRange("C4").setFormula("=C6-C5").setBackground(greenBg);
  sheet.getRange("D4").setFormula("=C4/$C$6").setBackground(greenBg).setNumberFormat("0%");
  
  sheet.getRange("B4:D6").setBorder(true, true, true, true, true, true);

  // --- 4. IDENTIFY COLUMNS ---
  const getCol = (name) => {
      const headers = sourceSheet.getRange(1, 1, 1, 15).getValues()[0];
      const idx = headers.indexOf(name);
      return idx >= 0 ? parseInt(idx + 1) : 0; 
  };

  const colFullName = getCol("Full Name"); 
  const colLinkName = getCol("Verkada Action Link Name"); 
  const colMeetingType = getCol("Meeting Type"); 
  const colInPerson = getCol("Verkada Action: In Person Meeting"); 

  console.log(`Cols Found: Name=${colFullName}, Link=${colLinkName}, Type=${colMeetingType}, IP=${colInPerson}`);

  // --- 5. BUILD PIVOT TABLE ---
  const anchorCell = sheet.getRange("B9");
  
  // LEGEND ROW (Row 8)
  sheet.getRange("B8").setValue("Blank = Untagged").setFontStyle("italic");
  sheet.getRange("D8").setValue("Verkada Action: In Person Meeting (0 = Virtual; 1 = In Person)")
       .setFontStyle("italic").setFontColor("#444444");

  const pivotTable = anchorCell.createPivotTable(sourceRange);

  if (colLinkName > 0 && colFullName > 0) {
      // *** FIX: RESTORED THE WORKING 2-ARGUMENT SYNTAX ***
      // We pass the column AND the function in the same command.
      const pivotValue = pivotTable.addPivotValue(colLinkName, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
      pivotValue.setDisplayName("Meetings"); // Changed to null to ignore "Meeting Type" showing up and creating 2 header rows

      const rowName = pivotTable.addRowGroup(colFullName);
      rowName.showTotals(true).sortDescending().sortBy(pivotValue, []);

      const rowType = pivotTable.addRowGroup(colMeetingType);
      rowType.showTotals(false).sortDescending().sortBy(pivotValue, []);

      const colGrp = pivotTable.addColumnGroup(colInPerson);
      colGrp.showTotals(true).sortAscending();
      // colGrp.showTotals(true).sortAscending().setDisplayName("Virtual or In Person");

      const criteria = SpreadsheetApp.newFilterCriteria().whenCellNotEmpty().build();
      pivotTable.addFilter(colFullName, criteria);

      // Filter out blank Meeting Types so empty rows don't appear - this filters out meetings that are 'blank'/untagged from the list
      // pivotTable.addFilter(colMeetingType, criteria);
  } else {
      sheet.getRange("B9").setValue("⚠️ Error: Source headers not found. Please re-run 'Repair All Sheets'.");
  }

  // --- 6. UNTAGGED LIST SECTION ---
  sheet.getRange("H2").setValue("List of untagged meetings in SFDC")
       .setFontWeight("bold")
       .setFontFamily("Poppins")
       .setFontSize(11)
       .setBackground("#283e4d")
       .setFontColor("white")
       .setHorizontalAlignment("center")
       .setBorder(true, true, true, true, true, true);

  sheet.getRange("H10").setFormula(`={'${sourceSheetName}'!A1:O1}`);
  sheet.getRange(10, 8, 1, 15)
       .setFontWeight("bold").setBackground("#283e4d").setFontColor("white").setHorizontalAlignment("center");

  const listFormula = `=filter('${sourceSheetName}'!A2:O,'${sourceSheetName}'!B2:B<>"",'${sourceSheetName}'!M2:M="")`;
  sheet.getRange("H11").setFormula(listFormula);

  sheet.setFrozenRows(10);
  
  // --- 7. STYLING & CLEANUP ---
  SpreadsheetApp.flush(); 
  
  const safeMaxRow = Math.max(sheet.getLastRow(), 500);

  // --- A. FONT STYLING (Poppins) ---
  sheet.getRange("B9:F10").setFontFamily("Poppins").setFontSize(11); // Pivot Headers
  sheet.getRange("H10:V10").setFontFamily("Poppins").setFontSize(11); // List Headers
  
  sheet.getRange(11, 2, safeMaxRow, 5).setFontFamily("Poppins").setFontSize(10); // Pivot Data
  sheet.getRange(11, 8, safeMaxRow, 15).setFontFamily("Poppins").setFontSize(10); // List Data

  // --- B. COLOR & ALIGNMENT ---
  sheet.getRange("B9:F10").setBackground("#283e4d") 
                          .setFontColor("white")
                          .setFontWeight("bold")
                          .setHorizontalAlignment("center");

  sheet.getRange(9, 2, safeMaxRow, 5).setHorizontalAlignment("center");
  // Force Column C (Meeting Type) to Left Align
  sheet.getRange(11, 3, safeMaxRow, 1).setHorizontalAlignment("left");

  // --- C. AUTO-RESIZE & BUFFERS ---
  SpreadsheetApp.flush(); 
  
  sheet.autoResizeColumns(2, 5); 
  for (let i = 2; i <= 6; i++) {
    sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 15);
  }

  sheet.autoResizeColumns(8, 15); 
  for (let i = 8; i <= 22; i++) {
    sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 35);
  }

  // --- D. MANUAL OVERRIDES ---
  if (sheet.getColumnWidth(5) < 120) { sheet.setColumnWidth(5, 120); }

  // --- E. HIGHLIGHT BLANK MEETING TYPES (YELLOW) & FIX GRAND TOTAL (LIGHT GREY 1) ---
  const lastPivotRow = sheet.getRange("B" + sheet.getMaxRows()).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();

  if (lastPivotRow > 10) {
    const typeRange = sheet.getRange(11, 3, lastPivotRow - 10, 1);
    const nameRange = sheet.getRange(11, 2, lastPivotRow - 10, 1);
    
    const typeValues = typeRange.getValues();
    const nameValues = nameRange.getValues();
    const typeBackgrounds = typeRange.getBackgrounds(); 
    
    let needsUpdate = false;

    for (let i = 0; i < typeValues.length; i++) {
      const rowName = nameValues[i][0];
      const cellValue = typeValues[i][0];

      if (rowName === "Grand Total") {
         // Force Grand Total row to "Light Grey 1"
         typeBackgrounds[i][0] = "#d9d9d9"; 
         needsUpdate = true;
      } else if (cellValue === "") {
         // Highlight empty untagged rows Yellow
         typeBackgrounds[i][0] = "#fff2cc"; 
         needsUpdate = true;
      }
    }

    if (needsUpdate) {
      typeRange.setBackgrounds(typeBackgrounds);
    }
  }

  ss.toast("SE Radar Main Dashboard has been repaired.");
}

// =================================================================
//      E N D   S E   R A D A R   M E E T I N G S   ( M A I N   D A S H )
// =================================================================

// #endregion
