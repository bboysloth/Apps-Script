/**
 * @file Meeting Tagger Tool v2.9.3
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
 * v2.9.3 --- added Hidden RADAR sheet
 */

// =================================================================
// #region 1. SPREADSHEET UI & MENU CREATION (v2.9.0 - Added Untagged Repair)
// =================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Meeting Tools')
    .addItem('Save and Refresh Untagged List', 'findEventsMissingTag')
    .addSeparator()
    .addItem('Apply Changes to Calendar', 'applyBatchChanges')
    .addSeparator()
    .addItem('Generate Visuals Dashboard', 'generateDashboard')
    .addSeparator()
    // Config Tools
    .addItem('Config: Update Quarter List', 'updateQuarterDropdown')
    .addItem('Config: Update Tag Dropdowns & Colors', 'updateTagDropdownsAndColors')
    .addItem('Config: Validate Tag Keywords', 'validateTagKeywords')
    .addSeparator()
    // New Admin Submenu
    .addSubMenu(ui.createMenu('Admin: Initialize/Repair Sheets')
        .addItem('Initialize ALL Sheets (Full Reset)', 'initializeAllSheets')
        .addSeparator()
        .addItem('Repair "Untagged Meetings" Sheet Only', 'repairUntaggedSheet')
        .addItem('Repair "Config" Sheet Only', 'repairConfigSheet')
        .addItem('Repair "Tags" Sheet Only', 'repairTagsSheet')
        .addItem('Repair "Filter Lists" Sheet Only', 'repairFilterSheet')
        .addItem('Repair "Instructions" Sheet Only', 'repairInstructionsSheet') //NEW v2.9.2
        .addItem('Repair "SE Radar Meetings" Sheet Only', 'repairRadarMainSheet') // NEW v2.9.5
        .addItem('Repair "All SE Radar Meetings CFQ" Sheet Only', 'repairRadarFilteredSheet')  // NEW v2.9.4
        .addItem('Repair "SFDC Radar Master" Sheet Only', 'repairRadarImportSheet')) // NEW v2.9.3
    .addToUi();
}
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

/**
 * Menu action to apply batch changes.
 */
function applyBatchChanges() {
  updateMeetingTag(null, false); 
}

/**
 * PRIMARY REFRESH ACTION:
 * 1. Saves Configuration (Updates Triggers based on MinutesInterval)
 * 2. Refreshes the Untagged Meetings List
 * * This consolidates "Save Config" and "Refresh List" into one safe action.
 */
function refreshUntaggedList() {
  try {
    SpreadsheetApp.flush(); // Ensure all user edits are saved before running

    // 1. UPDATE TRIGGERS (The "Save Config" part)
    // We do this every time to ensure the auto-refresh interval is always in sync.
    const file = SpreadsheetApp.getActiveSpreadsheet();
    const triggers = ScriptApp.getProjectTriggers();
    const minutes = parseInt(getConfig("MinutesInterval"), 10);

    // Delete ALL existing triggers for the refresher to ensure a clean slate
    triggers.forEach(trigger => {
      const handler = trigger.getHandlerFunction();
      if (handler === 'findEventsMissingTag') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // FIX: Only create triggers if MinutesInterval is actually > 0.
    // If 0, we create NO triggers (purely manual mode).
    if (minutes > 0) {
      // 1. Time-based trigger (e.g., every 30 mins)
      ScriptApp.newTrigger('findEventsMissingTag').timeBased().everyMinutes(minutes).create();
      
      // 2. OnOpen trigger (Only automatically refresh on load if Auto-Refresh is ENABLED)
      ScriptApp.newTrigger('findEventsMissingTag').forSpreadsheet(file).onOpen().create();
    }

    // 2. RUN THE LIST GENERATOR (The "Refresh" part)
    findEventsMissingTag();
    
    SpreadsheetApp.getActiveSpreadsheet().toast("Config saved & list refreshed.");

  } catch (e) {
    SpreadsheetApp.getUi().alert("Error during refresh: " + e.message);
  }
}
// #endregion

// =================================================================
// #region 3. CORE LOGIC FUNCTIONS (v4.2.0 - URL & Token Safety)
// =================================================================

function findEventsMissingTag(forcedUserEmail = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = ss; // Alias for compatibility
  try {
    // --- 1. CONFIGURATION ---
    const daysBack = parseInt(getConfig("DaysBack"), 10);
    const daysAhead = parseInt(getConfig("DaysAhead"), 10);
    const selectedQuarter = getConfig("QuarterOverride");
    const MAX_PATCH_DAYS_PAST = 500; 

    // Lists (Parsed for "Smart Matching")
    const ignorePhrasesRules = _parseMatchRule(_getFilterList("IgnorePhrases").join(","));
    
    // Exact lists remain literal
    const ignoreExactList = _getFilterList("IgnoreExactTitles");
    const ignoreFromEmailsList = _getFilterList("IgnoreFromEmails").map(s => s.toLowerCase());
    const ignoreToEmailsList = _getFilterList("IgnoreToEmails").map(s => s.toLowerCase());

    // Toggles
    const useIgnorePhrases = (getConfig("IgnorePhrases") || "false").toString().toLowerCase() === "true";
    const useIgnoreExact = (getConfig("IgnoreExactTitles") || "false").toString().toLowerCase() === "true";
    const useIgnoreFrom = (getConfig("IgnoreFromEmails") || "false").toString().toLowerCase() === "true";
    const useIgnoreTo = (getConfig("IgnoreToEmails") || "false").toString().toLowerCase() === "true";
    const useAutoSeLead = (getConfig("AutoSELeadKeywords") || "false").toString().toLowerCase() === "true";
    const useAutoSeCover = (getConfig("AutoSECoverageKeywords") || "false").toString().toLowerCase() === "true"; 
    const useAutoInPerson = (getConfig("AutoInPersonKeywords") || "false").toString().toLowerCase() === "true";

    // Settings
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

    // --- 2. PREPARE TAG DATA (WATERFALL LOGIC) ---
    const tagsSheet = file.getSheetByName("Tags");
    const tagsData = tagsSheet.getRange("B2:G" + tagsSheet.getLastRow()).getValues();
    
    const tagWaterfall = [];
    const tagMap = new Map(); 

    tagsData.forEach(row => {
      const tagName = row[1];
      const abbrTag = row[2] ? row[2].replace(/;/g, "").trim() : "";
      const keywordString = row[3];
      const checkTitle = row[4] === true;
      const checkDesc = row[5] === true;

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
    const userEmail = (forcedUserEmail || Session.getActiveUser().getEmail()).toLowerCase();
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
        
        // --- TEXT PREP (Title/Desc) ---
        const titleRaw = event.summary || "";
        
        // CLEANUP STEP 1: Title Search Preparation
        // 1. Remove URLs
        // 2. Remove "Long Words" (>50 chars) which are likely tokens/base64/broken links
        const titleSearch = titleRaw
             .replace(/(https?:\/\/[^\s]+)/g, '')
             .replace(/\S{50,}/g, '') 
             .toLowerCase();
        
        let descRaw = (event.description || "")
             .replace(/<br\s*\/?>/gi, '\n')
             .replace(/<p>/gi, '\n')
             .replace(/<\/p>/gi, '\n')
             .replace(/<(?!(?:https?|tel):)[^>]*>/gi, ''); 
        
        // CLEANUP STEP 2: Description Search Preparation
        // Same logic: Strip URLs, then strip massive text blobs
        const descSearch = descRaw
             .replace(/(https?:\/\/[^\s]+)/g, '')
             .replace(/\S{50,}/g, '')
             .toLowerCase();
             
        const fullSearch = titleSearch + " " + descSearch;

        if (event.start.date) {
            if (!includeAllDayEvents) continue; 
            const isMyEvent = event.organizer && (event.organizer.email.toLowerCase() === userEmail);
            
            const isTrip = _isMatch(fullSearch, nightsAwayRules);

            if (isMyEvent && isTrip) {
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
        
        // --- FILTER LIST LOGIC ---
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

        if (attendanceFilter === "Show Only Accepted/Maybe") {
            let myStatus = "none"; 
            if (organizerEmail === userEmail) {
                myStatus = "organizer";
            } else if (event.attendees) {
                const me = event.attendees.find(a => a.email.toLowerCase() === userEmail);
                if (me) myStatus = me.responseStatus;
            }
            if (!["accepted", "tentative", "organizer"].includes(myStatus)) continue; 
        }
        
        const hasMainTag = descRaw.includes(tagPrefix);
        
        // Parse Existing Tags
        let originalSeLead = seLeadTagText ? descRaw.includes(seLeadTagText) : false;
        let originalSeCover = seCoverTagText ? descRaw.includes(seCoverTagText) : false;
        let originalInPerson = false;
        let originalTag = "";
        let hasSETagInDescription = false;
        
        if (hasMainTag) {
            const lines = descRaw.split('\n');
            const tagLine = lines.find(line => line.trim().startsWith(tagPrefix));
            if (tagLine) {
                hasSETagInDescription = tagLine.includes("(SE)");
                let rawContent = tagLine.trim().substring(tagPrefix.length).trim();
                
                if (modifier && rawContent.includes(modifier)) {
                  originalInPerson = true;
                }
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
            const isVirtual = virtualExclusions.some(keyword => locationLower.includes(keyword));
            if (!isVirtual && locationLower.trim() !== "") {
                if (locationRequiresExternal && !hasExternalAttendee) {
                } else {
                    suggestedInPerson = true;
                }
            }
        }

        // --- CORE MATCHING LOGIC (WATERFALL) ---
        let tagFound = false;

        for (const rule of tagWaterfall) {
            let isMatch = false;
            // Uses the cleaned 'descSearch' (with URLs/Tokens removed)
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
        
        if (suggestedTag && useAutoInPerson && _isMatch(fullSearch, autoInPersonRules)) {
            suggestedInPerson = true;
        }

        const eventStartTime = new Date(event.start.dateTime || event.start.date);
        const rowData = [
            titleRaw, eventStartTime, organizerEmail, hasExternalAttendee, 
            suggestedSeCover, suggestedSeLead, suggestedInPerson, suggestedTag, 
            '', 
            `=HYPERLINK("${event.htmlLink}", "Open Event")`, 
            '', 
            originalSeCover, originalSeLead, originalInPerson, originalTag,
            durationInMinutes, 
            nightsCount 
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
    
    const headers = ["Title", "Start Time", "Created By", "External Attendees", "SE Cover", "SE Lead", "In Person", "Add Tag to Meeting", "Synced", "Event Link", "Last Refreshed", "Original Cover", "Original SE", "Original IP", "Original Tag", "Duration (min)", "Nights"];
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
      const editedCol = range.getColumn();
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
    const tagsRaw = tagsSheet.getRange("B2:G" + tagsSheet.getLastRow()).getValues();
    const validTags = tagsRaw.filter(row => row[2] && row[2].trim() !== "");
    const tagMap = new Map(validTags.map(row => [row[1], row[2]])); 

    const sheetData = sheet.getRange(2, 1, lastRow - 1, 17).getValues(); 
    
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
    let errorCount = 0; // Track permission errors
    
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

      const hasChanged = (currentSeCover !== originalSeCover) || (currentSeLead !== originalSeLead) || (currentInPerson !== originalInPerson) || (currentTag !== originalTag);

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

          if (currentSeLead && seLeadTag) {
             description += breakChar + breakChar + seLeadTag;
          }
          if (currentSeCover && seCoverTag) {
             description += breakChar + seCoverTag; 
          }
          
          if (currentTag && currentTag.trim() !== "" && currentTag !== "<Clear Tag>") {
            let abbreviatedTag = tagMap.get(currentTag); 
            if (!abbreviatedTag) { 
              continue; 
            }
            
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
          
          // --- SAFETY BLOCK: Prevent Service Account Crashes ---
          try {
              Calendar.Events.patch({ description: description }, calendarId, eventToUpdate.id, { sendUpdates: "none" });
              rowsToDelete.push(i + 2);
              updatedCount++;
          } catch (e) {
              console.log(`Permission Error patching event "${title}": ${e.message}`);
              errorCount++;
          }
          // ----------------------------------------------------
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
// #region 3a. REGEX PARSERS (v4.3.0 - Robust Spaces)
// =================================================================

/**
 * PARSER: Converts "vce(!coverage), demo" into rule objects.
 * FIX: Now supports spaces like "vce (!coverage)" without breaking.
 */
function _parseMatchRule(rawString) {
  if (!rawString) return [];
  const parts = rawString.split(',').map(s => s.trim()).filter(Boolean);
  const rules = [];

  parts.forEach(part => {
    // Regex: Capture Keyword (Group 1) and optional Exclusion (Group 2)
    // Handles cases like "vce" and "vce(!coverage)" and "vce (!coverage)"
    const match = part.match(/^(.+?)(?:\s*\(!([^)]+)\))?$/);
    
    if (match) {
        const keyword = match[1].trim().toLowerCase();
        let excludes = [];
        if (match[2]) {
            excludes = match[2].split('|').map(x => x.trim().toLowerCase()).filter(Boolean);
        }
        if (keyword) {
            rules.push({ keyword: keyword, excludes: excludes });
        }
    }
  });
  return rules;
}

/**
 * MATCHER: Checks text against parsed rules.
 * Returns TRUE if (Keyword is Found) AND (None of the Excludes are Found).
 */
function _isMatch(text, rules) {
  if (!text || !rules || rules.length === 0) return false;
  for (const rule of rules) {
    // 1. Check Keyword
    if (text.includes(rule.keyword)) {
        // 2. Check Exclusions
        let isExcluded = false;
        if (rule.excludes.length > 0) {
            for (const excl of rule.excludes) {
                if (text.includes(excl)) {
                    isExcluded = true;
                    break;
                }
            }
        }
        // If keyword matched and NOT excluded, we have a winner.
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
  const headers = listSheet.getRange('A1:1').getValues()[0];
  const colIndex = headers.findIndex(h => h && h.toString().trim() === listName);
  if (colIndex === -1) return []; 
  const columnData = listSheet.getRange(2, colIndex + 1, listSheet.getLastRow() - 1, 1).getValues();
  return columnData.flat().filter(String).map(value => value.toString().trim());
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
    if (!silentMode) {
      ss.toast("Updating tag dropdowns and colors...");
    }

    const lastTagRow = tagsSheet.getLastRow();
    // Read Tag Name (Col C) and Colors
    const tagRawValues = tagsSheet.getRange("C2:C" + lastTagRow).getValues();
    const tagRawBackgrounds = tagsSheet.getRange("C2:C" + lastTagRow).getBackgrounds();
    
    const tagNames = [["<Clear Tag>"]]; 
    const tagColors = [["#ffffff"]]; 
    
    for (let i = 0; i < tagRawValues.length; i++) {
        if (tagRawValues[i][0] && tagRawValues[i][0].trim() !== "") {
            tagNames.push([tagRawValues[i][0]]);
            tagColors.push([tagRawBackgrounds[i][0]]);
        }
    }

    // FIX: Target Column 8 (H)
    const dropdownRange = mainSheet.getRange(2, 8, mainSheet.getMaxRows() - 1, 1);
    dropdownRange.clearDataValidations();

    const helperSheetName = "_TagListHelper";
    let helperSheet = ss.getSheetByName(helperSheetName);
    if (!helperSheet) { helperSheet = ss.insertSheet(helperSheetName).hideSheet(); } else { helperSheet.clear(); }

    if (tagNames.length > 0) {
        const helperRange = helperSheet.getRange(1, 1, tagNames.length, 1);
        helperRange.setValues(tagNames);
        helperRange.setBackgrounds(tagColors);
        dropdownRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(helperRange, true).setAllowInvalid(false).build());
    }

    mainSheet.clearConditionalFormatRules();
    let rules = [];
    
    // FIX: Target Column 8 (H) for Rules
    const dropdownRangeForRules = mainSheet.getRange(2, 8, mainSheet.getMaxRows() - 1, 1);
    tagNames.forEach((nameArr, i) => {
        if (nameArr[0] === "<Clear Tag>") {
          rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("<Clear Tag>").setBackground(tagColors[i][0]).setRanges([dropdownRangeForRules]).build());
        } else {
          rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(nameArr[0]).setBackground(tagColors[i][0]).setRanges([dropdownRangeForRules]).build());
        }
    });

    // Synced Status is now Column 9 (I)
    const syncedRange = mainSheet.getRange(2, 9, mainSheet.getMaxRows() - 1, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("TRUE").setBackground('#d9ead3').setFontColor('#38761d').setRanges([syncedRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("FALSE").setBackground('#f4cccc').setFontColor('#cc0000').setRanges([syncedRange]).build());
    
    mainSheet.setConditionalFormatRules(rules);

    if (!silentMode) {
      ss.toast("Tag dropdowns and colors updated.");
      try {
        const flagCell = tagsSheet.getRange("H1");
        flagCell.setValue("SAVED")
                .setBackground("#d9ead3") // Green
                .setFontWeight("bold")
                .setHorizontalAlignment("center");
      } catch (e) { }
    }
}

function validateTagKeywords() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tagsSheet = ss.getSheetByName("Tags");
    if (!tagsSheet) { ss.toast("Error: 'Tags' sheet not found."); return; }
    ss.toast("Validating keywords...");

    const tagsData = tagsSheet.getRange("C2:E" + tagsSheet.getLastRow()).getValues();
    const tracker = new Map(), duplicates = new Map();

    tagsData.forEach(row => {
        if (row[0] && row[2]) {
            row[2].split(',').map(k => k.trim().toLowerCase()).filter(String).forEach(k => {
                if (tracker.has(k) && tracker.get(k) !== row[0]) {
                    if (!duplicates.has(k)) duplicates.set(k, [tracker.get(k)]);
                    duplicates.get(k).push(row[0]);
                } else { tracker.set(k, row[0]); }
            });
        }
    });

    if (duplicates.size === 0) { ss.toast("Success! All keywords are unique."); }
    else {
        let msg = "Warning: Duplicate keywords found!\n";
        for (const [k, t] of duplicates.entries()) { msg += `\n- "${k}": used by ${[...new Set(t)].join(', ')}`; }
        SpreadsheetApp.getUi().alert("Validation Failed", msg, SpreadsheetApp.getUi().ButtonSet.OK);
    }
}

function onEditTagsSheet(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() === "Tags") {
    try {
      const flagCell = sheet.getRange("H1");
      const flagValue = flagCell.getValue();

      if (e.range.getA1Notation() === 'H1' && flagValue === "CHANGES DETECTED") {
        updateTagDropdownsAndColors(false); 
        return; 
      }
      
      if (e.range.getRow() > 1 && e.range.getColumn() >= 2 && e.range.getColumn() <= 7) { 
          flagCell.setValue("CHANGES DETECTED")
                  .setBackground("#f4cccc") 
                  .setFontWeight("bold")
                  .setHorizontalAlignment("center")
                  .setWrap(true);
      }
    } catch (err) { }
  }
}
// #endregion

// =================================================================
// #region 6. VISUALS DASHBOARD FUNCTIONS (v3.29.0 - Toggle Discrepancy Table)
// =================================================================

function _formatMinutesToHours(totalMinutes) {
  if (totalMinutes === 0) return "0min";
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  
  let result = "";
  if (hours > 0) result += `${hours}hr`;
  if (minutes > 0) result += ` ${minutes}min`;
  return result.trim();
}

function generateDashboard(showDiscrepancyTable = false) {
  // --- CONFIGURATION FLAGS ---
  // OLD LOGIC WITHOUT BRIDGE SCRIPT 
  //const SHOW_DISCREPANCY_TABLE = true; // Set to FALSE to hide the discrepancy table
  // ---------------------------



  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Untagged Meetings");
  // Check for singular or plural sheet name for Radar
  const radarSheet = ss.getSheetByName("All SE Radar Meetings CFQ") || ss.getSheetByName("SFDC - Tagged SE Radar Meetings"); 
  const visualsSheetName = "Visuals";
  
  let visualsSheet = ss.getSheetByName(visualsSheetName);
  if (!visualsSheet) { 
      visualsSheet = ss.insertSheet(visualsSheetName); 
  } else { 
      // 1. Remove Charts
      visualsSheet.getCharts().forEach(c => visualsSheet.removeChart(c));
      
      // 2. Clear Content & Formats
      visualsSheet.clear(); 
      
      // 3. NUCLEAR UN-GROUPING (The Fix)
      // First, force ALL rows to be visible (unhide everything)
      const maxRows = visualsSheet.getMaxRows();
      visualsSheet.showRows(1, maxRows);
      
      // Second, loop to remove nested groups (up to 8 levels deep)
      // We grab the entire Column A to target all rows.
      const fullRange = visualsSheet.getRange(1, 1, maxRows, 1);
      
      for (let i = 0; i < 8; i++) {
        try {
          // Flatten groups one level at a time
          fullRange.shiftRowGroupDepth(-1);
        } catch(e) {
          // If this errors, it usually means there are no groups left to remove, 
          // so we are done.
          break; 
        }
      }
  }
  
  visualsSheet.showColumns(1, visualsSheet.getMaxColumns());
  visualsSheet.getRange("A1:Z1000").setFontFamily("Poppins").setFontSize(10);
  visualsSheet.setFrozenRows(6); 
  
  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert("No data found in Untagged Meetings sheet."); return; }
  
  // --- 1. PROCESS CALENDAR DATA ---
  const data = dataSheet.getRange(2, 1, lastRow - 1, 17).getValues();
  
  let total = data.length, external = 0, seLead = 0, seCover = 0, inPerson = 0, totalMinutes = 0, totalNights = 0;
  
  const groupedStats = new Map();
  const getBaseName = (name) => name.replace(/\s*\(.*?\)\s*$/, "").trim();

  data.forEach(row => {
    const isExt = row[3] === true;
    const isCover = row[4] === true;
    const isLead = row[5] === true;
    const isIP = row[6] === true;
    const tagName = row[7];
    const duration = Number(row[15]) || 0; 
    const nights = Number(row[16]) || 0;

    if (isExt) external++;
    if (isCover) seCover++;
    if (isLead) seLead++;
    if (isIP) inPerson++;
    totalMinutes += duration;
    totalNights += nights;
    
    if (tagName && tagName.trim() !== "") { 
        const baseName = getBaseName(tagName);
        
        if (!groupedStats.has(baseName)) {
            groupedStats.set(baseName, { 
                variants: new Map(), 
                radar: { count: 0, ip: 0, lead: 0 },
                agg: { count: 0, time: 0, ip: 0, lead: 0, cover: 0, ext: 0, nights: 0 }
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
                  const rTagStr = rTag.toString().trim();
                  const rBase = getBaseName(rTagStr);
                  
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

  const bg = "#283e4d";
  const fg = "white";

  // Scorecards Row 1
  const row1Cards = [
    { title: "Total Meetings", val: fTotal, col: 1 },
    { title: "External Meetings", val: fExternal, col: 3 },
    { title: "Internal Meetings", val: fInternal, col: 5 }
  ];
  row1Cards.forEach(card => {
    visualsSheet.getRange(3, card.col, 1, 2).merge().setValue(card.title).setFontWeight("bold").setHorizontalAlignment("center").setBackground(bg).setFontColor(fg).setBorder(true, true, true, true, true, true);
    visualsSheet.getRange(4, card.col, 1, 2).merge().setFormula(card.val).setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  });

  // Total Time
  visualsSheet.getRange(3, 7, 1, 2).merge().setValue("Total Blocked Mtg Time").setFontWeight("bold").setHorizontalAlignment("center").setBackground(bg).setFontColor(fg).setBorder(true, true, true, true, true, true);
  visualsSheet.getRange(4, 7, 1, 2).merge().setFormula(`=TEXT(SUMIF(O11:O, FALSE, AC11:AC)/1440, "[h]""hr ""mm""min""")`).setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);

  // Scorecards Row 2
  const row2Cards = [
    { title: "In Person Mtgs", val: fInPerson, col: 1 },
    { title: "SE Lead Mtgs", val: fSeLead, col: 3 },
    { title: "SE Coverage Mtgs", val: fSeCover, col: 5 }, 
    { title: "Nights Away", val: fNights, col: 7 }   
  ];
  row2Cards.forEach(card => {
    visualsSheet.getRange(5, card.col, 1, 2).merge().setValue(card.title).setFontWeight("bold").setHorizontalAlignment("center").setBackground(bg).setFontColor(fg).setBorder(true, true, true, true, true, true);
    visualsSheet.getRange(6, card.col, 1, 2).merge().setFormula(card.val).setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
  });

  const scRanges = [visualsSheet.getRange(4, 1, 1, 8), visualsSheet.getRange(6, 1, 1, 8)];
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
          data.agg.count,
          data.agg.ip,
          data.agg.lead,
          data.agg.cover,
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

      hiddenMetricRows.push([data.agg.ext, data.agg.nights, data.agg.time]);
      
      parentRowsIndices.push(currentRowIndex); 
      const parentRowIdx = currentRowIndex;
      currentRowIndex++;

      // Child Rows
      if (isMulti) {
          const variants = [...data.variants.entries()].sort((a, b) => b[1].count - a[1].count);
          variants.forEach(([tagName, vStats]) => {
              displayRows.push([
                  "   ↳ " + tagName,
                  vStats.count,
                  vStats.ip,
                  vStats.lead,
                  vStats.cover,
                  _formatMinutesToHours(vStats.time)
              ]);
              hiddenMetricRows.push([0, 0, 0]);
              currentRowIndex++;
          });
          groupsToCreate.push({ start: parentRowIdx + 1, num: variants.length });
      }
  });

  // --- 5. WRITE MAIN TABLE ---
  const mainHeaderRow = 9;
  const mainStartRow = 11;
  const mainEndRow = mainStartRow + displayRows.length;
  
  visualsSheet.getRange(mainHeaderRow, 8, 1, 8).merge().setValue("From Untagged Meetings List and Calendar")
              .setFontWeight("bold").setFontColor(fg).setBackground("#3d9fd2") 
              .setHorizontalAlignment("center").setFontSize(11).setBorder(true, true, true, true, true, true);

  const headers = ["Tag Name", "Total", "IP", "Lead", "Cov", "Block Time", "%", "Hide"];
  visualsSheet.getRange(mainHeaderRow + 1, 8, 1, 8).setValues([headers]).setFontWeight("bold").setBackground(bg).setFontColor(fg);
  visualsSheet.getRange(mainHeaderRow + 1, 8, 1, 5).setHorizontalAlignment("center");
  visualsSheet.getRange(mainHeaderRow + 1, 15, 1, 1).setHorizontalAlignment("center");
  visualsSheet.getRange(mainHeaderRow + 1, 13, 1, 2).setHorizontalAlignment("right"); 

  if (displayRows.length > 0) {
    visualsSheet.getRange(mainStartRow, 8, displayRows.length, 6).setValues(displayRows);
    visualsSheet.getRange(mainStartRow, 27, hiddenMetricRows.length, 3).setValues(hiddenMetricRows).setNumberFormat("0");

    // Checkboxes
    visualsSheet.getRange(mainStartRow, 15, visualsSheet.getMaxRows() - mainStartRow, 1).removeCheckboxes();
    parentRowsIndices.forEach(idx => visualsSheet.getRange(idx, 15).insertCheckboxes());
    
    // Formulas
    const pctFormulas = [];
    for (let i = mainStartRow; i < mainEndRow; i++) {
        pctFormulas.push([`=IF(O${i}="", "-", IF(O${i}=TRUE, "-", I${i}/SUMIF($O$11:$O,FALSE,$I$11:$I)))`]);
    }
    visualsSheet.getRange(mainStartRow, 14, displayRows.length, 1).setFormulas(pctFormulas).setNumberFormat("0.0%");

    // Styling
    visualsSheet.getRange(mainStartRow, 9, displayRows.length, 4).setNumberFormat("0").setHorizontalAlignment("center");
    visualsSheet.getRange(mainStartRow, 13, displayRows.length, 2).setHorizontalAlignment("right");
    
    parentRowsIndices.forEach(idx => {
        visualsSheet.getRange(idx, 8).setFontWeight("bold");
    });

    visualsSheet.getRange(mainStartRow, 9, displayRows.length, 1).setBackground("#f3f3f3");
    visualsSheet.getRange(mainStartRow, 13, displayRows.length, 2).setBackground("#f3f3f3");

    visualsSheet.getRange(mainHeaderRow + 1, 8, displayRows.length + 1, 8).setBorder(true, true, true, true, true, true);
  }

  // --- 6. WRITE COMPARISON TABLE ---
  SpreadsheetApp.flush(); 
  
  // Prepare the base rules (Scorecards) for application later
  let finalFormatRules = [...scorecardRules]; 

    if (showDiscrepancyTable) { //SET TO TRUE TO BUILD, OTHERWISE IS FALSE
      const compHeaderRow = mainEndRow + 4; 
      const compStartRow = compHeaderRow + 2;
      const compLen = comparisonRows.length;
      
      if (compLen > 0) {
          // 6.1 Title Bar
          const titleRange = visualsSheet.getRange(compHeaderRow, 8, 1, 10);
          titleRange.breakApart(); 
          titleRange.merge();
          titleRange.setValue("Discrepancy Check (Calendar vs Radar)");
          titleRange.setFontWeight("bold").setFontColor("white").setBackground("#6aa84f")
                    .setHorizontalAlignment("center").setFontSize(11)
                    .setBorder(true, true, true, true, true, true);

          // 6.2 Column Headers
          const compHeaders = ["Tag Name", "Cal Total", "Radar Total", "Diff", "Cal IP", "Radar IP", "Diff", "Cal Lead", "Radar Lead", "Diff"];
          visualsSheet.getRange(compHeaderRow + 1, 8, 1, 10).setValues([compHeaders])
                      .setFontWeight("bold").setBackground("#efefef").setFontColor("black")
                      .setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);

          // 6.3 Data
          visualsSheet.getRange(compStartRow, 8, compLen, 10).setValues(comparisonRows);
          
          visualsSheet.getRange(compStartRow, 9, compLen, 9).setHorizontalAlignment("center").setNumberFormat("0");
          visualsSheet.getRange(compStartRow, 8, compLen, 1).setFontWeight("bold"); 
          
          visualsSheet.getRange(compStartRow, 9, compLen, 1).setBackground("#f3f3f3");
          visualsSheet.getRange(compStartRow, 12, compLen, 1).setBackground("#f3f3f3");
          visualsSheet.getRange(compStartRow, 15, compLen, 1).setBackground("#f3f3f3");
          
          visualsSheet.getRange(compHeaderRow + 1, 8, compLen + 1, 10).setBorder(true, true, true, true, true, true);

          // 6.4 Conditional Rules (Append only if table is shown)
          const diffRanges = [
              visualsSheet.getRange(compStartRow, 11, compLen, 1),
              visualsSheet.getRange(compStartRow, 14, compLen, 1),
              visualsSheet.getRange(compStartRow, 17, compLen, 1)
          ];
          
          const diffGreenRule = SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setFontColor("#38761d").setBold(true).setBackground("#d9ead3").setRanges(diffRanges).build();
          const diffRedRule = SpreadsheetApp.newConditionalFormatRule().whenNumberNotEqualTo(0).whenCellNotEmpty().setFontColor("#cc0000").setBold(true).setBackground("#f4cccc").setRanges(diffRanges).build();

          // Add these new rules to our master list
          finalFormatRules = finalFormatRules.concat([diffGreenRule, diffRedRule]);
      }
  }

  // APPLY ALL RULES (Scorecards + Discrepancy if enabled)
  visualsSheet.setConditionalFormatRules(finalFormatRules);

  // --- 7. APPLY GROUPING ---
  groupsToCreate.forEach(grp => {
      try {
          visualsSheet.getRange(grp.start, 1, grp.num, 1).shiftRowGroupDepth(1);
          const parentRow = grp.start - 1;
          visualsSheet.getRange(parentRow, 8, 1, 8).setBackground("#efefef");
      } catch(e) {}
  });

  // --- CHART SOURCE (Fix: Dynamic Formula) ---
  visualsSheet.getRange("AD10").setValue("Chart Source");
  
  // FIX: Dynamic filter that ignores hidden rows (O=FALSE) AND child rows (Starts with space)
  visualsSheet.getRange("AD11").setFormula(`=FILTER(H11:I, O11:O=FALSE, LEFT(H11:H, 1) <> " ")`);
  
  SpreadsheetApp.flush(); 

  // Since we use a formula, we don't know exact rows, so we grab a large range for the chart.
  if (displayRows.length > 0) {
    // Dynamic range AD11:AE
    const chartRange = visualsSheet.getRange("AD11:AE" + (11 + displayRows.length)); 
    
    const pieChart = visualsSheet.newChart().setChartType(Charts.ChartType.PIE).addRange(chartRange)
      .setOption('title', 'Meeting Tag Breakdown (Aggregated)')
      .setOption('pieSliceText', 'percentage')
      .setOption('is3D', true)
      .setOption('colors', ['#283e4d', '#3d9fd2', '#757475', '#34545e', '#959ea7', '#546e7a', '#78909c', '#63c0f2', '#1f4e6a', '#de6662', '#c45551', '#e68a87'])
      .setOption('titleTextStyle', { fontName: 'Poppins', fontSize: 20, bold: true })
      .setOption('pieSliceTextStyle', { fontName: 'Poppins', color: 'white' })
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH)
      
      .setOption('legend', { position: 'right', textStyle: { fontName: 'Poppins', fontSize: 13 } }) 
      .setOption('chartArea', { left: '5%', top: '10%', width: '70%', height: '80%' }) 
      
      .setOption('height', 600)
      .setOption('width', 700) 
      .setPosition(9, 1, 0, 0).build(); 
    visualsSheet.insertChart(pieChart);
  } else { visualsSheet.getRange("A10").setValue("No tagged meetings found."); }

  const selectedQuarter = getConfig("QuarterOverride");
  let dateString = `Last ${getConfig("DaysBack")} Days`;
  if (selectedQuarter && selectedQuarter !== "None") { dateString = selectedQuarter; }
  visualsSheet.getRange("A1").setValue(`Dashboard for: ${dateString} (Generated: ${new Date().toLocaleString()})`).setFontWeight("bold").setFontSize(12);

  visualsSheet.setColumnWidth(8, 250); 
  visualsSheet.setColumnWidth(9, 60); 
  visualsSheet.setColumnWidth(10, 50); 
  visualsSheet.setColumnWidth(11, 50); 
  visualsSheet.setColumnWidth(12, 50); 
  visualsSheet.setColumnWidth(13, 90); 
  visualsSheet.setColumnWidth(14, 60); 
  visualsSheet.setColumnWidth(15, 50); 
  
  for (let c = 9; c <= 17; c++) {
      visualsSheet.setColumnWidth(c, 85); 
  }
  
  [1, 2, 3, 5, 6, 7].forEach(c => visualsSheet.setColumnWidth(c, 100));
  
  visualsSheet.hideColumns(27, 5); 

  ss.setActiveSheet(visualsSheet);
}
// #endregion

// =================================================================
// #region 7. INITIALIZATION & SCAFFOLDING (v3.14.0 - Optional SE Name Dropdown)
// =================================================================

function initializeAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  repairInstructionsSheet();
  repairRadarMainSheet(); // <--- NEW: Auto-generates the 'SE Radar Meetings' Dashboard Tab 
  repairRadarImportSheet(); // <--- NEW: Auto-generates the hidden 'SFDC - Radar Data' Tab
  repairRadarFilteredSheet();  // <--- NEW: The Visible Filtered Tab
  repairUntaggedSheet(); 
  repairConfigSheet();
  repairFilterSheet();
  repairTagsSheet();
  
  repairSheetStructure(); 

  const visualsSheet = ss.getSheetByName("Visuals");
  if (visualsSheet) visualsSheet.clear();

  ss.toast("Full system initialization complete.");
}

function repairSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ["Instructions","Untagged Meetings", "Config", "Filter Lists", "Tags", "Visuals", "SE Radar Meetings", "All SE Radar Meetings CFQ"];
  
  sheets.forEach((name, index) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) { sheet = ss.insertSheet(name); }
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(index + 1);
  });
}

// --- NEW FUNCTION ---
function repairInstructionsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Instructions");
  if (!sheet) { sheet = ss.insertSheet("Instructions", 0); }

  sheet.clear();
  sheet.setHiddenGridlines(true); // Cleaner look

  // Title Header
  sheet.getRange("A1:B1").merge().setValue("SE Radar Tool - User Guide")
       .setFontWeight("bold").setFontSize(14)
       .setBackground("#283e4d").setFontColor("white")
       .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setFrozenRows(1);

  const instructions = [
    ["Creating an Untagged Meeting List:"],
    ["- Configure your desired settings and filters on the 'Config' tab."],
    ["- You can Update/Refresh your list via the 'Meeting Tools > Save and Refresh Untagged List' Menu Dropdown."],
    ["- Add keywords to the 'Filter Lists' and 'Tags' sheets for auto-tag suggestions. You will need to Refresh for those terms to be included"],
    ["------ You can also implement 'Exclusion' logic to a search term to exclude keyword combinations from an auto-tag suggestion"],
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
  const sectionHeaderIndices = [0, 10, 17]; 
  sectionHeaderIndices.forEach(idx => {
      sheet.getRange(startRow + idx, 1).setFontWeight("bold").setFontColor("#1155cc").setFontSize(11); 
  });

  // Column Width
  sheet.setColumnWidth(1, 800);
  
  ss.toast("Instructions sheet repaired.");
}


function repairUntaggedSheet() {
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

  for (let i = 1; i <= maxCols; i++) {
      sheet.setColumnWidth(i, 100);
  }

  ss.toast("Untagged Meetings sheet fully reset.");
}

function repairConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName("Config");
  if (!configSheet) configSheet = ss.insertSheet("Config");
  
  configSheet.clear(); 
  
  // Base Configuration Data
  let configData = [];

  // ==========================================================================================
  // --- [OPTIONAL] SFDC SE NAME DROPDOWN BLOCK (START) ---
  // Delete everything between START and END if moving to a tool without the SFDC sheet.
  // This inserts "SE Name" at the very top (Row 2).
  
  configData.push({ 
      key: "SE Name", 
      val: "", // Default value is empty, user selects from dropdown
      desc: "Select your name to see SFDC meetings on the 'SE Radar' Tabs", 
      type: "se_name_dropdown" // Special type handled below
  });
  
  // --- [OPTIONAL] SFDC SE NAME DROPDOWN BLOCK (END) ---
  // ==========================================================================================

  // Standard Configuration Data
  const standardConfig = [
    { key: "MinutesInterval", val: 30, desc: "How frequently (30 minutes max) the Untagged Meetings tab automatically refreshes. 0 = No auto-refresh", type: "number" },
    { key: "QuarterOverride", val: "None", desc: "When selected, start/end dates align with Verkada's FY dates. API restricts edits to about 2 quarters in the past\n**Selecting 'None' allows you to to use 'DaysBack' and 'DaysAhead' logic", type: "quarter_dropdown" },
    { key: "DaysBack", val: 30, desc: "How far back from today (in days) to look for untagged meetings (ignored if QuarterOverride selected)", type: "number" },
    { key: "DaysAhead", val: 1, desc: "How far ahead from today (in days) to look for untagged meetings (ignored if QuarterOverride selected)", type: "number" },
    
    { key: "IgnorePhrases", val: true, desc: `CHECKED: Meetings are ignored if they CONTAIN words/phrases located in the IgnorePhrases Filter List.\nUNCHECKED: Ignore this Filter List`, type: "checkbox" },
    { key: "IgnoreExactTitles", val: true, desc: `CHECKED: Exclude Event Titles found in the IgnoreExactTitles Filter List from results\nUNCHECKED: Ignore this Filter List`, type: "checkbox" },
    { key: "IgnoreFromEmails", val: true, desc: `CHECKED: Ignore meetings from specific email addresses located in the IgnoreFromEmails Filter List\nUNCHECKED: Ignore this Filter List`, type: "checkbox" },
    { key: "IgnoreToEmails", val: true, desc: `CHECKED: Ignore meetings sent to specific email addresses located in the IgnoreToEmails Filter List\nUNCHECKED: Ignore this Filter List`, type: "checkbox" },
    
    { key: "IncludeAllDayEvents", val: true, desc: "CHECKED : The script shows all-day events on the \"Untagged Meetings\" list (use with 'NightsAwayKeywords')\nUNCHECKED: The script ignores any all-day events.", type: "checkbox" },
    { key: "NightsAwayKeywords", val: true, desc: `Used with 'IncludeAllDayEvents' filter for ACTUAL Travel Events. Manage keyword suggestions from the NightsAway Filter List`, type: "checkbox" },
    
    { key: "LocationSetsInPerson", val: true, desc: "CHECKED: Suggests an 'In Person' tag on any meeting containing a physical address/location\nUNCHECKED: Does NOT suggest an 'In Person' tag on meetings containing a location", type: "checkbox" },
    { key: "LocationRequiresExternal", val: true, desc: "CHECKED: Location-based 'In Person' suggestions ONLY occur if an External Attendee is present (Filters out internal conference rooms).\nUNCHECKED: Any physical location triggers an 'In Person' suggestion.", type: "checkbox" },
    
    { key: "AutoInPersonKeywords", val: true, desc: `Scan the full title + description against your AutoInPersonKeywords Filter list. If it finds a match, it will check the 'In Person' checkbox as a suggested calendar event tag MODIFIER. Comma-delimited`, type: "checkbox" },
    { key: "AutoSELeadKeywords", val: true, desc: `Scan the full title + description against your AutoSELeadKeywords Filter list. If it finds a match, it will check the 'SE Lead' checkbox as a suggested calendar event tag edit. Comma-delimited`, type: "checkbox" },
    { key: "AutoSECoverageKeywords", val: true, desc: `Scan the full title + description against your AutoSECoverageKeywords Filter list. If it finds a match, it will check the 'SE Cover' checkbox as a suggested calendar event tag edit. Comma-delimited`, type: "checkbox" }, 
    
    { key: "FilterForMissingSETags", val: false, desc: "CHECKED: Show only meetings with no tag, and meetings that have an AE/SFDC implemented tags.\nUNCHECKED: HIDES meetings with any tag (AE/SFDC OR SE Tags)\n**Results will populate with YELLOW if the meeting is missing a ('SE') Tag, and RED if the meeting has NO Tags at all.", type: "checkbox" },
    { key: "ShowTaggedMeetings", val: true, desc: "CHECKED: Displays all tagged meetings (with GREEN) allowing for editing past tagged events\noverrides 'FilterForMissingSETags' behavior to hide \"(SE)\" tagged meetings", type: "checkbox" },
    { key: "ExternalAttendees", val: true, desc: "CHECKED: SHOWS only Meetings with External Attendees\nUNCHECKED: SHOWS Meetings with and without External Attendees (Internal Meetings)", type: "checkbox" },
    
    { key: "InPersonModifier", val: "-P", desc: "This modifier is added to a meeting type if you are In Person instead of joining Virtually.\nThis should not need to be changed but has been put here incase logic changes with Meeting Tags in the future", type: "text" },
    { key: "SELeadTagText", val: "Verkada SE Lead;", desc: "This denotes the text added to a meeting type if you lead the call.\nThis should not need to be changed but has been put here incase logic changes with Meeting Tags in the future", type: "text" },
    { key: "SECoverageTagText", val: "Verkada SE Coverage;", desc: "Text added to description when 'SE Cover' is checked.", type: "text" },
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
  configSheet.getRange("C:C").setWrap(true);
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
    // --- [OPTIONAL] SE NAME DROPDOWN LOGIC ---
    else if (type === "se_name_dropdown") {
      // Safely check if the source sheet exists
      const sfdcSheet = ss.getSheetByName("SFDC - Tagged SE Radar Meetings") || ss.getSheetByName("SFDC - Tagged SE Radar Meeting");
      
      if (sfdcSheet) {
          const lastRow = sfdcSheet.getLastRow();
          // Ensure there is data to create a dropdown from
          if (lastRow >= 3) {
              // 1. ATTEMPT TO SORT THE SOURCE SHEET
              // Since the list is >500 items, we must verify the sheet itself is sorted.
              // We sort the entire data range to keep rows aligned.
              try {
                  const fullDataRange = sfdcSheet.getRange(3, 1, lastRow - 2, sfdcSheet.getLastColumn());
                  fullDataRange.sort({column: 1, ascending: true});
              } catch (e) {
                  // Note: If Column A is populated by a spreadsheet FORMULA (e.g., =IMPORTRANGE),
                  // this sort command will fail or do nothing. 
                  // In that case, you must wrap your spreadsheet formula in =SORT(...) directly.
              }

              // 2. APPLY VALIDATION "FROM RANGE"
              // This bypasses the 500-item limit.
              const range = sfdcSheet.getRange("A3:A");
              const rule = SpreadsheetApp.newDataValidation()
                  .requireValueInRange(range)
                  .build();
              
              cell.setDataValidation(rule);
          }
      }
    }
    // -----------------------------------------
    else {
      cell.clearDataValidations();
    }
  }
  ss.toast("Config sheet repaired.");
}

function repairFilterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let filterSheet = ss.getSheetByName("Filter Lists");
  if (!filterSheet) filterSheet = ss.insertSheet("Filter Lists");
  
  filterSheet.clear();

  const syncFormula = `=TRANSPOSE(QUERY(Config!A:A, "SELECT A WHERE A IS NOT NULL AND (A CONTAINS 'Keywords' OR A CONTAINS 'Ignore' OR A CONTAINS 'Location') AND A != 'AttendanceFilter' AND A != 'FilterForMissingSETags' AND A != 'LocationSetsInPerson' AND A != 'LocationRequiresExternal' AND A != 'SE Name'", 0))`;
  filterSheet.getRange("A1").setFormula(syncFormula);
  filterSheet.getRange("1:1").setFontWeight("bold").setBackground("#283e4d").setFontColor("white")
              .setFontFamily("Poppins").setVerticalAlignment("middle");
  filterSheet.setFrozenRows(1);

  const filterDefaults = {
    "IgnorePhrases":        [""], 
    "IgnoreExactTitles":    [""], 
    "IgnoreFromEmails":     [""], 
    "IgnoreToEmails":       [""], 
    "NightsAwayKeywords":   ["hotel", "trip", "stay"],  
    "AutoInPersonKeywords": ["sitewalk", "site walk", "vce"], 
    "AutoSELeadKeywords":   [""],
    "AutoSECoverageKeywords": [""] 
  };

  const defaultColumns = [
    filterDefaults["IgnorePhrases"],
    filterDefaults["IgnoreExactTitles"],
    filterDefaults["IgnoreFromEmails"],
    filterDefaults["IgnoreToEmails"],
    filterDefaults["NightsAwayKeywords"],
    filterDefaults["AutoInPersonKeywords"],
    filterDefaults["AutoSELeadKeywords"],
    filterDefaults["AutoSECoverageKeywords"]
  ];

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
  
  ss.toast("Filter Lists sheet repaired.");
}

function repairTagsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tagsSheet = ss.getSheetByName("Tags");
  if (!tagsSheet) tagsSheet = ss.insertSheet("Tags");
  
  tagsSheet.clear(); 

  const tagHeaders = [["Demofy Tag Section", "Tag and Color (For Dropdown)", "Abbreviated Tag (For Calendar)", "Keywords (auto-tag suggestions)", "Titles", "Descriptions", "SAVED"]];
  tagsSheet.getRange("B1:H1").setValues(tagHeaders)
            .setFontWeight("bold").setBackground("#283e4d").setFontColor("white")
            .setHorizontalAlignment("center").setVerticalAlignment("middle");
  
  const tagsLayout = [
    ["Sales Cycle", "", "", "", "", "", ""], 
    ["", "Dedicated Discovery Call (AE)", "DS;", "", "", "", ""],
    ["", "Demo (AE)", "DM;", "", "", "", ""],
    ["", "Best Practice (AE)", "BP;", "", "", "", ""],
    ["", "Project Scoping (AE)", "PS;", "", "", "", ""],
    ["", "Site Walk (AE)", "SW;", "", "", "", ""],
    ["", "Pricing / Proposal (AE)", "PN;", "", "", "", ""],
    ["", "Executive Alignment / Board Meeting (AE)", "EX;", "", "", "", ""],
    
    ["Customer Success", "", "", "", "", "", ""], 
    ["", "Existing Customer Check-In / CBR (AE)", "OC;", "", "", "", ""],
    ["", "Existing Customer Support (AE)", "SP;", "", "", "", ""],

    ["Solutions Engineering", "", "", "", "", "", ""], 
    ["", "Demo (SE)", "DM (SE);", "demo", true, "", ""],
    ["", "Best Practice (SE)", "BP (SE);", "best practice", true, "", ""],
    ["", "Trial Setup/Config (SE)", "TC (SE);", "trial, setup, config, configuration", true, "", ""],
    ["", "Site Walk (SE)", "SW (SE);", "sitewalk, site walk", true, "", ""],
    ["", "Floorplans (SE)", "FP (SE);", "floorplan", true, "", ""],
    ["", "VCE Training (SE)", "VT (SE);", "vce", true, "", ""],
    ["", "Partner Onboarding / Training (SE)", "PT (SE);", "", "", "", ""],
    ["", "Industry Conference / Trade Show (SE)", "TS (SE);", "", "", "", ""],
    ["", "Verkada-Sponsored Event (SE)", "VM (SE);", "", "", "", ""],
    ["", "Existing Customer Support (SE)", "SP (SE);", "", "", "", ""],
    ["", "Existing Customer Check-In / CBR (SE)", "OC (SE);", "", "", "", ""],
    ["", "AE Enablement (SE)", "AE (SE);", "", "", "", ""],
    ["", "RFP / RFI / Security Questionnaire (SE)", "RQ (SE);", "", "", "", ""],
    ["", "EBC (SE)", "EB (SE);", "", "", "", ""],

    ["Partner", "", "", "", "", "", ""], 
    ["", "Account Mapping (PARTNER)", "AM;", "", "", "", ""],
    ["", "Deal Related Discussion (PARTNER)", "DR;", "", "", "", ""],
    ["", "Partner Onboarding / Training (PARTNER)", "PT;", "", "", "", ""],
    ["", "Channel Co-Marketing (PARTNER)", "CM;", "", "", "", ""],
    ["", "Partner Demo Certification (PARTNER)", "DC;", "", "", "", ""],

    ["Other", "", "", "", "", "", ""], 
    ["", "Recruiting (OTHER)", "RT;", "", "", "", ""],
    ["", "VCE Training (OTHER)", "VC;", "", "", "", ""],
    ["", "Other (OTHER)", "OT;", "", "", "", ""],

    ["Personal", "", "", "", "", "", ""], 
    ["", "Viper Team", "VPR (SE);", "Viper", true, "", ""],
    ["", "Verkada Internal Team Discussions", "INT (SE);", "", "", "", ""],
    ["", "Travel Time", "TRVL (SE);", "Flight, Layover, Drive to, Commute", true, "", ""],
    ["", "Overnight", "SLEEP (SE);", "Stay:, Stay At", true, "", ""]
  ];

  const dataRange = tagsSheet.getRange(2, 2, tagsLayout.length, 7);
  dataRange.setValues(tagsLayout);
  
  dataRange.setFontColor("black").setBackground("white");

  tagsSheet.getRange("A:Z").setFontFamily("Poppins").setVerticalAlignment("middle");
  tagsSheet.setFrozenRows(1);
  tagsSheet.setColumnWidth(2, 200); 
  tagsSheet.setColumnWidth(3, 300); 
  tagsSheet.setColumnWidth(5, 300); 
  tagsSheet.setColumnWidth(6, 60); 
  tagsSheet.setColumnWidth(7, 100); 

  // --- DYNAMIC FORMATTING LOGIC ---
  const sectionPalette = [
    "#efefef", // Grey (Sales)
    "#ead1dc", // Purple (CS)
    "#d9ead3", // Green (SE)
    "#cfe2f3", // Blue (Partner)
    "#fce5cd", // Orange (Other)
    "#d9d2e9", // Lavender (Future 1)
    "#fff2cc", // Yellow (Future 2)
    "#f4cccc"  // Red (Future 3)
  ];
  
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
          tagsSheet.getRange(rowNum, 2, 1, 7).setBackground("#283e4d").setFontColor("white").setFontWeight("bold");
      } else if (tagName !== "") {
          tagsSheet.getRange(rowNum, 5, 1, 3).setBackground("#f3f3f3");
          const checkRange = tagsSheet.getRange(rowNum, 6, 1, 2);
          checkRange.insertCheckboxes();
          
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

  tagsSheet.getRange("H1").setValue("SAVED")
           .setBackground("#d9ead3").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
           
  const tagsFooterStart = tagsLayout.length + 3; 
  const tagsFooterText = [
      ["1. SAVE CHANGES: 'Meeting Tools > Update Tag Dropdowns and Colors'"],
      ["2. COLUMN 'H' WILL TELL YOU IF THE CHANGES MADE HERE WILL SHOW ON YOUR UNTAGGED MEETING LIST"],
      ["3. CHANGES MADE HERE WILL NEED YOU TO SAVE THE LIST OR YOU WILL GET 'DATA VALIDATION' ERRORS"]
  ];

  tagsSheet.getRange(tagsFooterStart, 2, tagsFooterText.length, 1).setValues(tagsFooterText);
  tagsSheet.getRange(tagsFooterStart, 2, tagsFooterText.length, 1)
           .setFontColor("#cc0000").setFontWeight("bold");
  
  updateTagDropdownsAndColors(true);
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
