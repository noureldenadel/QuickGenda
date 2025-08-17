// SmartAgenda_Creator.jsx (v1.0 - Initial Professional Release)
// Date: August 17, 2025
// Author: noorr
// 
// ** CORE FEATURES **
// ==================
// 
// CSV PROCESSING & VALIDATION:
// • Dynamic CSV column detection with comma/semicolon auto-detection
// • UTF-8 encoding support with BOM handling
// • Robust CSV parsing with quoted field support
// • Smart header validation and field mapping
// • Session grouping with topic aggregation
// • Comprehensive error handling and user feedback
// 
// TEMPLATE MANAGEMENT:
// • Automatic template document detection and inheritance
// • Document settings preservation (page size, margins, binding, facing pages)
// • Master page duplication and management
// • Intelligent placeholder detection system
// • Multi-layout template support (table vs. independent)
// 
// CHAIRPERSON LAYOUT SYSTEM:
// • Dual-mode layout: Inline text or grid-based positioning
// • Advanced grid configuration (rows/columns, spacing, units)
// • Smart horizontal centering with vertical position retention
// • Prototype group management and cloning
// • Custom separator options (comma, line breaks)
// • Double-pipe delimiter support (||) for chairperson separation
// 
// TOPIC MANAGEMENT:
// • Flexible topic layouts: Table-based or independent row positioning
// • Dynamic table creation with configurable headers
// • Independent row layout with precise spacing control
// • Topic field consistency (topicTime, topicTitle, topicSpeaker)
// • Group-aware positioning and cloning
// 
// ADVANCED TEXT PROCESSING:
// • Universal line break replacement system
// • Configurable character-to-linebreak conversion
// • Support for all CSV fields (session, chairperson, topic data)
// • Frame-level and table cell text processing
// • Overset text detection and reporting
// 
// IMAGE AUTOMATION SYSTEM:
// • Intelligent chairperson image placement
// • Multi-format support (JPG, PNG, TIF, PSD)
// • Advanced name variation generation and matching
// • Avatar and flag image detection
// • Configurable image fitting options
// • Comprehensive search attempt logging
// • Professional title stripping (Prof., Dr., MD, PhD)
// • Name normalization and variation generation
// 
// UNIFIED SETTINGS MANAGEMENT:
// • Centralized settings panel with tabbed navigation
// • Complete settings export/import functionality
// • JSON-based configuration persistence
// • Cross-session settings retention
// • Template-aware option enabling/disabling
// • Real-time UI synchronization
// 
// PROFESSIONAL REPORTING:
// • Comprehensive import success/failure reporting
// • Overset text detection and alerting
// • Image placement success tracking
// • Detailed error logging and troubleshooting info
// • Session summary with topic counts
// • Missing image file tracking with search attempts
// 
// ENTERPRISE-GRADE ERROR HANDLING:
// • Graceful degradation on missing template elements
// • Comprehensive try-catch blocks throughout
// • User-friendly error messages with actionable guidance
// • File I/O error handling and recovery
// • Memory management and cleanup
// 
// TECHNICAL SPECIFICATIONS:
// • InDesign CS6+ compatibility
// • ExtendScript-optimized JSON handling
// • Unicode text support
// • Cross-platform file path handling
// • Efficient DOM manipulation
// • Memory-conscious object management
// 
// ** SUPPORTED CSV STRUCTURE **
// ============================
// Required Fields:
// • "Session Title" - Primary session identifier
// 
// Optional Fields:
// • "Session Time" - Session timing information
// • "Session No" - Session numbering/identification
// • "Chairpersons" - Session chairpersons (|| separated)
// • "Time" - Individual topic timing
// • "Topic Title" - Topic description
// • "Speaker" - Topic presenter information
// 
// ** TEMPLATE REQUIREMENTS **
// ===========================
// 
// Essential Labels:
// • sessionTitle - Session title text frame
// • sessionTime - Session time text frame
// • sessionNo - Session number text frame
// • chairpersons - Chairpersons text frame or group prototype
// • topicsTable - Topics table frame (for table layout)
// 
// Topic Layout Labels (Independent mode):
// • topicTime - Topic time text frame
// • topicTitle - Topic title text frame  
// • topicSpeaker - Topic speaker text frame
// 
// Image Automation Labels (Optional):
// • chairAvatar - Avatar image frame within chairperson groups
// • chairFlag - Flag image frame within chairperson groups
// 
// ** VERSION HISTORY **
// ====================
// v1.0 (August 17, 2025) - Initial Professional Release
//   • Complete rewrite and optimization
//   • Unified settings management system
//   • Advanced image automation
//   • Comprehensive error handling
//   • Professional reporting system
//   • Enterprise-grade template management

#target indesign

function main() {
    if (app.documents.length === 0) {
        alert("Error: No document is open.\nPlease open your template file first.");
        return;
    }
    var templateDoc = app.activeDocument;

    // --- Create a new document from the template ---
    var doc = app.documents.add();

    // Copy document settings from template to new document BEFORE any page operations
    copyDocumentSettings(templateDoc, doc);

    // Duplicate template pages to the new document (this will be our master template)
    templateDoc.pages.everyItem().duplicate(LocationOptions.AFTER, doc.pages.lastItem());

    // Now remove the original default page (after we have template pages)
    if (doc.pages.length > 1) {
        doc.pages[0].remove();
    }

    // --- Detect layout capabilities of the template ---
    var hasIndependentTopicLayout = detectIndependentTopicSetup(doc.pages[0]);
    var hasTableTopicLayout = !!findItemByLabel(doc.pages[0], "topicsTable");

    // Pick CSV
    var csvFile = File.openDialog("Select your CSV (comma- or semicolon-delimited)", "*.csv");
    if (!csvFile) {
        doc.close(SaveOptions.NO);
        return;
    }

    // Parse CSV to sessions with dynamic column detection
    var parseResult = parseCSVDynamic(csvFile);
    if (!parseResult || parseResult.sessions.length === 0) {
        alert("No sessions found in CSV.");
        doc.close(SaveOptions.NO);
        return;
    }

    var sessionsData = parseResult.sessions;
    var availableFields = parseResult.availableFields;

    // Show user what fields were detected
    var fieldsList = [];
    for (var field in availableFields) {
        if (availableFields[field]) fieldsList.push(field);
    }
    alert("CSV loaded successfully!\n\nDetected fields: " + fieldsList.join(", ") + 
          "\n\nSessions found: " + sessionsData.length);

    // Get all user options from unified settings panel
    var allOptions = getUnifiedSettingsPanel(hasIndependentTopicLayout, hasTableTopicLayout, availableFields);
    if (!allOptions) { doc.close(SaveOptions.NO); return; }

    var templatePage = doc.pages[0]; // This is now our master template page
    for (var i = 0; i < sessionsData.length; i++) {
        // Always duplicate the template page for each session
        var currentPage = templatePage.duplicate(LocationOptions.AFTER, doc.pages.lastItem());
        fillPagePlaceholders(currentPage, sessionsData[i], allOptions.chairOptions, allOptions.topicOptions, allOptions.lineBreakOptions, availableFields);
    }

    // Remove the master template page after creating all session pages
    templatePage.remove();

    // Ask if user wants to export a report
    if (confirm("Agenda created successfully!\n" + sessionsData.length + " sessions were imported.\n\nWould you like to export a report?")) {
        var reportFile = exportReport(sessionsData);
        if (reportFile) {
            alert("Report saved to:\n" + reportFile.fsName + 
                  "\n\nNote: You can now export and import all your layout settings\n" +
                  "using the centralized buttons in the main settings panel.");
        }
    }
}

/* ---------------- LAYOUT DETECTION ---------------- */

function detectIndependentTopicSetup(page) {
    // Check for topic frames with new naming convention
    var timeFrame = findItemByLabel(page, "topicTime");
    var topicFrame = findItemByLabel(page, "topicTitle");
    var speakerFrame = findItemByLabel(page, "topicSpeaker");
    
    return !!(timeFrame && topicFrame && speakerFrame);
}

/* ---------------- DYNAMIC CSV PARSING ---------------- */

function parseCSVDynamic(file) {
    file.encoding = "UTF-8";
    file.open("r");
    var txt = file.read();
    file.close();

    if (!txt || txt === "") return null;

    // Strip UTF-8 BOM if present
    if (txt.length > 0 && txt.charCodeAt(0) === 0xFEFF) {
        txt = txt.substring(1);
    }
    // Normalize line endings
    txt = txt.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

    var lines = txt.split("\n");
    // Skip leading empty lines
    var start = 0;
    while (start < lines.length && trimES(lines[start]) === "") start++;
    if (start >= lines.length) return null;

    // Detect delimiter from header line (prefer comma)
    var headerLine = lines[start];
    var commaCount = countChar(headerLine, ',');
    var semiCount  = countChar(headerLine, ';');
    var delim = (commaCount >= semiCount) ? ',' : ';';

    var headers = parseSeparatedLine(headerLine, delim);
    for (var h = 0; h < headers.length; h++) headers[h] = trimES(headers[h]);

    // Define possible fields (all optional now)
    var possibleFields = {
        "Session Title": false,
        "Session Time": false,
        "Session No": false,
        "Chairpersons": false,  // Changed from Chairperson(s) to Chairpersons
        "Time": false,
        "Topic Title": false,
        "Speaker": false
    };

    // Check which fields are available in the CSV
    var idx = {};
    for (var field in possibleFields) {
        var foundIdx = indexOfExact(headers, field);
        if (foundIdx !== -1) {
            possibleFields[field] = true;
            idx[field] = foundIdx;
        }
    }

    // Require at least Session Title for sessions to work
    if (!possibleFields["Session Title"]) {
        alert("CSV Header Error: The CSV file must contain at least a 'Session Title' column.\n\nAvailable headers (" + (delim === ',' ? "comma" : "semicolon") + "-delimited):\n" + headers.join(" | "));
        return null;
    }

    var sessions = [];
    var currentSession = null;
    var lastTitle = null;

    for (var i = start + 1; i < lines.length; i++) {
        var raw = lines[i];
        if (!raw || trimES(raw) === "") continue;

        var row = parseSeparatedLine(raw, delim);
        while (row.length < headers.length) row.push("");

        var sessionTitle = possibleFields["Session Title"] ? trimES(row[idx["Session Title"]]) : "";
        var sessionTime = possibleFields["Session Time"] ? trimES(row[idx["Session Time"]]) : "";
        var sessionNo = possibleFields["Session No"] ? trimES(row[idx["Session No"]]) : "";
        var chairpersons = possibleFields["Chairpersons"] ? trimES(row[idx["Chairpersons"]]) : "";  // Changed from Chairperson(s) to Chairpersons
        var time = possibleFields["Time"] ? trimES(row[idx["Time"]]) : "";
        var topicTitle = possibleFields["Topic Title"] ? trimES(row[idx["Topic Title"]]) : "";
        var speaker = possibleFields["Speaker"] ? trimES(row[idx["Speaker"]]) : "";

        if (currentSession === null || sessionTitle !== lastTitle) {
            currentSession = { 
                title: sessionTitle, 
                time: sessionTime,
                no: sessionNo,
                chairs: chairpersons, 
                topics: [] 
            };
            sessions.push(currentSession);
            lastTitle = sessionTitle;
        } else {
            // Update session fields if they weren't filled initially
            if (!currentSession.time && sessionTime) currentSession.time = sessionTime;
            if (!currentSession.no && sessionNo) currentSession.no = sessionNo;
            if (!currentSession.chairs && chairpersons) currentSession.chairs = chairpersons;
        }

        // Add topic if any topic field has content
        if (topicTitle !== "" || time !== "" || speaker !== "") {
            currentSession.topics.push({ time: time, title: topicTitle, speaker: speaker });
        }
    }

    return {
        sessions: sessions,
        availableFields: possibleFields
    };
}

function parseSeparatedLine(line, delim) {
    var out = [];
    var cur = "";
    var inQuotes = false;

    for (var i = 0; i < line.length; i++) {
        var ch = line.charAt(i);

        if (ch === '"') {
            if (inQuotes && i + 1 < line.length && line.charAt(i + 1) === '"') {
                cur += '"'; i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (ch === delim && !inQuotes) {
            out.push(cur); cur = "";
        } else {
            cur += ch;
        }
    }
    out.push(cur);
    return out;
}

function countChar(s, ch) { var c = 0; for (var i = 0; i < s.length; i++) if (s.charAt(i) === ch) c++; return c; }
function indexOfExact(arr, str) { for (var i = 0; i < arr.length; i++) if (arr[i] === str) return i; return -1; }
function trimES(s) { if (s === null || s === undefined) return ""; return s.replace(/^\s+|\s+$/g, ""); }

/* ---------------- UNIFIED SETTINGS PANEL ---------------- */

function getUnifiedSettingsPanel(hasIndependentTopicLayout, hasTableTopicLayout, availableFields) {
    var defaultSettings = getDefaultSettings();
    var dlg = new Window('dialog', 'SmartAgenda - Configuration');
    dlg.alignChildren = 'fill';
    dlg.preferredSize.width = 600;
    dlg.preferredSize.height = 700;
    
    // Navigation tabs
    var tabGroup = dlg.add('tabbedpanel');
    tabGroup.preferredSize.width = 580;
    tabGroup.preferredSize.height = 600;
    
    // Create tab panels
    var chairTab = tabGroup.add('tab', undefined, 'Chairpersons');
    var topicTab = tabGroup.add('tab', undefined, 'Topics');
    var lineBreakTab = tabGroup.add('tab', undefined, 'Line Breaks');
    
    // ===== CHAIRPERSONS TAB =====
    setupChairpersonsTab(chairTab, defaultSettings);
    
    // ===== TOPICS TAB =====
    setupTopicsTab(topicTab, defaultSettings, hasIndependentTopicLayout, hasTableTopicLayout);
    
    // ===== LINE BREAKS TAB =====
    setupLineBreaksTab(lineBreakTab, defaultSettings, availableFields);
    
    // ===== CENTRALIZED SETTINGS MANAGEMENT =====
    var pnlSettings = dlg.add('panel', undefined, 'Settings Management');
    pnlSettings.orientation = 'row';
    pnlSettings.alignChildren = 'center';
    pnlSettings.margins = [15, 15, 15, 15];
    
    var btnImport = pnlSettings.add('button', undefined, 'Import All Settings');
    var btnExport = pnlSettings.add('button', undefined, 'Export All Settings');
    btnImport.preferredSize.width = 150;
    btnExport.preferredSize.width = 150;
    
    // Import all settings button handler
    btnImport.onClick = function() {
        var importedSettings = importAllSettings();
        if (importedSettings) {
            // Update all tabs with imported settings
            updateChairpersonsTabFromSettings(chairTab, importedSettings.chairOptions || {});
            updateTopicsTabFromSettings(topicTab, importedSettings.topicOptions || {});
            updateLineBreaksTabFromSettings(lineBreakTab, importedSettings.lineBreakOptions || {});
            alert('All settings imported successfully!');
        }
    };
    
    // Export all settings button handler
    btnExport.onClick = function() {
        var allSettings = collectAllSettings(chairTab, topicTab, lineBreakTab);
        if (exportAllSettings(allSettings)) {
            alert('All settings exported successfully!');
        }
    };
    
    // Dialog buttons with credit
    var grpBtns = dlg.add('group');
    grpBtns.orientation = 'row';
    grpBtns.alignment = 'fill';

    // Add credit on the left side
    var txtCredit = grpBtns.add('statictext', undefined, 'noorr @2025');
    txtCredit.graphics.font = ScriptUI.newFont(txtCredit.graphics.font.name, ScriptUI.FontStyle.ITALIC, 10);
    txtCredit.graphics.foregroundColor = txtCredit.graphics.newPen(txtCredit.graphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5], 1);

    // Add invisible spacer with fixed width to push buttons right
    var spacer = grpBtns.add('statictext', undefined, '');
    spacer.preferredSize.width = 350; // Adjust this number to control spacing

    // Add buttons directly to main group (not sub-group)
    var btnOK = grpBtns.add('button', undefined, 'OK');
    var btnCancel = grpBtns.add('button', undefined, 'Cancel', { name: 'cancel' });

    // Set button widths for consistency
    btnOK.preferredSize.width = 80;
    btnCancel.preferredSize.width = 80;
    
    if (dlg.show() !== 1) return null;
    
    // Collect all settings and return
    return collectAllSettings(chairTab, topicTab, lineBreakTab);
}

function setupChairpersonsTab(tab, defaultSettings) {
    tab.alignChildren = 'fill';
    tab.margins = [15, 15, 15, 15];
    
    // Layout mode selection
    var pnlDesc = tab.add('panel', undefined, 'Layout Selection');
    pnlDesc.margins = [15, 15, 15, 15];
    pnlDesc.alignChildren = 'fill';
    
    var grpMode = pnlDesc.add('group'); 
    grpMode.orientation = 'row';
    grpMode.spacing = 10;
    grpMode.add('statictext', undefined, 'Layout:').preferredSize.width = 80;
    var ddMode = grpMode.add('dropdownlist', undefined, ['Inline (single frame)', 'Separate frames grid']);
    ddMode.selection = defaultSettings.chairMode === 'grid' ? 1 : 0;
    ddMode.preferredSize.width = 200;
    
    // Inline options
    var pnlInline = tab.add('panel', undefined, 'Inline Layout Configuration');
    pnlInline.orientation = 'column'; 
    pnlInline.alignChildren = 'left';
    pnlInline.margins = [15, 15, 15, 15];
    
    var grpInlineOptions = pnlInline.add('group');
    grpInlineOptions.orientation = 'row';
    grpInlineOptions.spacing = 15;
    var rbComma = grpInlineOptions.add('radiobutton', undefined, 'Comma separated');
    var rbLines = grpInlineOptions.add('radiobutton', undefined, 'Line breaks');
    rbComma.value = defaultSettings.chairInlineSeparator !== 'linebreak';
    rbLines.value = defaultSettings.chairInlineSeparator === 'linebreak';
    
    // Grid options
    var pnlGrid = tab.add('panel', undefined, 'Grid Layout Configuration');
    pnlGrid.orientation = 'column'; 
    pnlGrid.alignChildren = 'left';
    pnlGrid.margins = [15, 15, 15, 15];
    
    var grpOrder = pnlGrid.add('group'); 
    grpOrder.orientation = 'row';
    grpOrder.spacing = 10;
    grpOrder.add('statictext', undefined, 'Fill order:').preferredSize.width = 80;
    var ddOrder = grpOrder.add('dropdownlist', undefined, ['Row-first', 'Column-first']);
    ddOrder.selection = defaultSettings.chairOrder === 'col' ? 1 : 0;
    ddOrder.preferredSize.width = 150;
    
    var grpDims = pnlGrid.add('group'); 
    grpDims.orientation = 'row';
    grpDims.spacing = 10;
    grpDims.add('statictext', undefined, 'Columns:').preferredSize.width = 80;
    var etCols = grpDims.add('edittext', undefined, defaultSettings.chairColumns || '2'); 
    etCols.characters = 4;
    grpDims.add('statictext', undefined, 'Rows:').preferredSize.width = 50;
    var etRows = grpDims.add('edittext', undefined, defaultSettings.chairRows || '2'); 
    etRows.characters = 4;
    
    var grpSpace = pnlGrid.add('group'); 
    grpSpace.orientation = 'row';
    grpSpace.spacing = 10;
    grpSpace.add('statictext', undefined, 'Col spacing:').preferredSize.width = 80;
    var etColSpace = grpSpace.add('edittext', undefined, defaultSettings.chairColSpacing || '8'); 
    etColSpace.characters = 6;
    grpSpace.add('statictext', undefined, 'Row spacing:').preferredSize.width = 80;
    var etRowSpace = grpSpace.add('edittext', undefined, defaultSettings.chairRowSpacing || '4'); 
    etRowSpace.characters = 6;
    
    var grpUnits = pnlGrid.add('group'); 
    grpUnits.orientation = 'row';
    grpUnits.spacing = 10;
    grpUnits.add('statictext', undefined, 'Units:').preferredSize.width = 80;
    var ddUnits = grpUnits.add('dropdownlist', undefined, ['pt', 'mm', 'cm', 'px']);
    ddUnits.selection = defaultSettings.chairUnits ? indexOfExact(['pt', 'mm', 'cm', 'px'], defaultSettings.chairUnits) : 0;
    
    var grpCenter = pnlGrid.add('group');
    grpCenter.orientation = 'row';
    grpCenter.spacing = 10;
    var cbCenter = grpCenter.add('checkbox', undefined, 'Center grid horizontally');
    cbCenter.value = defaultSettings.chairCenterGrid || false;
    
    // Image automation controls - moved back to grid configuration as they're related
    var grpImageEnable = pnlGrid.add('group');
    grpImageEnable.orientation = 'row';
    grpImageEnable.spacing = 10;
    var cbEnableImages = grpImageEnable.add('checkbox', undefined, 'Enable automatic image placement');
    cbEnableImages.value = defaultSettings.chairEnableImages || false;
    
    var grpFolderPath = pnlGrid.add('group');
    grpFolderPath.orientation = 'row';
    grpFolderPath.spacing = 10;
    grpFolderPath.add('statictext', undefined, 'Image folder:').preferredSize.width = 80;
    var etImageFolder = grpFolderPath.add('edittext', undefined, defaultSettings.chairImageFolder || '');
    etImageFolder.preferredSize.width = 200;
    var btnBrowseFolder = grpFolderPath.add('button', undefined, 'Browse...');
    btnBrowseFolder.preferredSize.width = 80;
    
    var grpImageFitting = pnlGrid.add('group');
    grpImageFitting.orientation = 'row';
    grpImageFitting.spacing = 10;
    grpImageFitting.add('statictext', undefined, 'Image fitting:').preferredSize.width = 80;
    var ddImageFitting = grpImageFitting.add('dropdownlist', undefined, ['Use Template Default', 'Fill Frame', 'Fit Proportionally', 'Fit Content to Frame', 'Center Content']);
    ddImageFitting.selection = defaultSettings.chairImageFitting ? indexOfExact(['Use Template Default', 'Fill Frame', 'Fit Proportionally', 'Fit Content to Frame', 'Center Content'], defaultSettings.chairImageFitting) : 0;
    ddImageFitting.preferredSize.width = 150;
    
    var txtImageHelp = pnlGrid.add('statictext', undefined, 
        'Places avatar (name.jpg) and flag (flag-name.jpg) images into labeled frames.\n' +
        'Requires chairAvatar and/or chairFlag labels within chairperson groups.');
    txtImageHelp.graphics.font = ScriptUI.newFont(txtImageHelp.graphics.font.name, ScriptUI.FontStyle.ITALIC, 10);
    
    // Browse folder handler
    btnBrowseFolder.onClick = function() {
        var folder = Folder.selectDialog('Select folder containing chairperson images');
        if (folder) {
            etImageFolder.text = folder.fsName;
        }
    };
    
    // Enable/disable image controls
    function syncImageUI() {
        grpFolderPath.enabled = cbEnableImages.value;
        grpImageFitting.enabled = cbEnableImages.value;
        txtImageHelp.enabled = cbEnableImages.value;
    }
    
    cbEnableImages.onClick = syncImageUI;
    
    // CSV format help
    var pnlNote = tab.add('panel', undefined, 'CSV Format Help');
    pnlNote.margins = [15, 15, 15, 15];
    pnlNote.add('statictext', undefined, 'Use double pipe (||) to separate chairpersons in your CSV file.\n' +
                                         'Example: "John Smith||Jane Doe||Robert Johnson"');
    
    // Store references for later access
    tab.ddMode = ddMode;
    tab.rbComma = rbComma;
    tab.rbLines = rbLines;
    tab.ddOrder = ddOrder;
    tab.etCols = etCols;
    tab.etRows = etRows;
    tab.etColSpace = etColSpace;
    tab.etRowSpace = etRowSpace;
    tab.ddUnits = ddUnits;
    tab.cbCenter = cbCenter;
    tab.cbEnableImages = cbEnableImages;
    tab.etImageFolder = etImageFolder;
    tab.ddImageFitting = ddImageFitting;
    
    // Make syncUI accessible from outside
    tab.syncUI = function() {
        var gridMode = ddMode.selection.index === 1;
        pnlInline.enabled = !gridMode;
        pnlGrid.enabled = gridMode;
        if (gridMode) { 
            var rowFirst = ddOrder.selection.index === 0; 
            etCols.enabled = rowFirst; 
            etRows.enabled = !rowFirst; 
        }
        // Sync image controls - only available in grid mode
        syncImageUI();
    };
    
    ddMode.onChange = tab.syncUI; 
    ddOrder.onChange = tab.syncUI;
    tab.syncUI(); // Initial sync
}

function setupTopicsTab(tab, defaultSettings, hasIndependentLayout, hasTableLayout) {
    tab.alignChildren = 'fill';
    tab.margins = [15, 15, 15, 15];
    
    // Layout mode selection
    var pnlDesc = tab.add('panel', undefined, 'Layout Selection');
    pnlDesc.margins = [15, 15, 15, 15];
    pnlDesc.alignChildren = 'fill';
    
    var pnlMode = tab.add('panel', undefined, 'Layout Mode');
    pnlMode.alignChildren = 'left';
    pnlMode.margins = [15, 15, 15, 15];
    
    var grpRadios = pnlMode.add('group');
    grpRadios.orientation = 'row';
    grpRadios.spacing = 20;
    var rbTable = grpRadios.add('radiobutton', undefined, 'Table Layout');
    var rbIndependent = grpRadios.add('radiobutton', undefined, 'Independent Row Layout');

    // Table layout options
    var pnlTable = tab.add('panel', undefined, 'Table Layout Configuration');
    pnlTable.alignChildren = 'left';
    pnlTable.margins = [15, 15, 15, 15];
    
    var grpHeader = pnlTable.add('group');
    grpHeader.orientation = 'row';
    grpHeader.spacing = 10;
    var cbHeader = grpHeader.add('checkbox', undefined, 'Include Table Header Row');
    cbHeader.value = defaultSettings.topicIncludeHeader !== false;

    // Independent layout options
    var pnlIndependent = tab.add('panel', undefined, 'Independent Row Configuration');
    pnlIndependent.alignChildren = 'left';
    pnlIndependent.margins = [15, 15, 15, 15];
    
    var grpSpace = pnlIndependent.add('group');
    grpSpace.orientation = 'row';
    grpSpace.spacing = 10;
    grpSpace.add('statictext', undefined, 'Vertical spacing:').preferredSize.width = 100;
    var etSpacing = grpSpace.add('edittext', undefined, defaultSettings.topicVerticalSpacing || '4'); 
    etSpacing.characters = 6;
    var ddSpaceUnits = grpSpace.add('dropdownlist', undefined, ['pt', 'mm', 'cm', 'px']);
    ddSpaceUnits.selection = defaultSettings.topicUnits ? indexOfExact(['pt', 'mm', 'cm', 'px'], defaultSettings.topicUnits) : 0;

    // Auto-select & disable options based on template detection
    rbTable.enabled = hasTableLayout;
    rbIndependent.enabled = hasIndependentLayout;

    // Use default settings if available, otherwise use template detection
    if (defaultSettings.topicMode) {
        if (defaultSettings.topicMode === 'table' && hasTableLayout) {
            rbTable.value = true;
        } else if (defaultSettings.topicMode === 'independent' && hasIndependentLayout) {
            rbIndependent.value = true;
        } else if (hasTableLayout) {
            rbTable.value = true;
        } else if (hasIndependentLayout) {
            rbIndependent.value = true;
        }
    } else {
        if (hasTableLayout) {
            rbTable.value = true;
        } else if (hasIndependentLayout) {
            rbIndependent.value = true;
        }
    }
    
    // Store references for later access
    tab.rbTable = rbTable;
    tab.rbIndependent = rbIndependent;
    tab.cbHeader = cbHeader;
    tab.etSpacing = etSpacing;
    tab.ddSpaceUnits = ddSpaceUnits;
    
    function syncUI() {
        var tableMode = rbTable.value;
        pnlTable.enabled = tableMode;
        pnlIndependent.enabled = !tableMode;
    }
    
    rbTable.onClick = syncUI;
    rbIndependent.onClick = syncUI;
    syncUI();
}

function setupLineBreaksTab(tab, defaultSettings, availableFields) {
    tab.alignChildren = 'fill';
    tab.margins = [15, 15, 15, 15];
    
    // Main description panel
    var pnlDesc = tab.add('panel', undefined, 'Line Break Configuration');
    pnlDesc.margins = [15, 15, 15, 15];
    
    // Session-level fields
    var pnlSession = tab.add('panel', undefined, 'Session Fields');
    pnlSession.alignChildren = 'fill';
    pnlSession.margins = [15, 15, 15, 15];
    
    var sessionControls = {};
    var sessionFields = ['SessionTitle', 'SessionTime', 'SessionNo'];
    var sessionDisplayNames = ['Session Title', 'Session Time', 'Session No'];
    var sessionCsvMapping = {
        'SessionTitle': 'Session Title',
        'SessionTime': 'Session Time',
        'SessionNo': 'Session No'
    };
    
    for (var i = 0; i < sessionFields.length; i++) {
        var field = sessionFields[i];
        var displayName = sessionDisplayNames[i];
        var csvField = sessionCsvMapping[field];
        
        var grp = pnlSession.add('group');
        grp.orientation = 'row';
        grp.alignChildren = 'left';
        grp.spacing = 10;
        
        var cb = grp.add('checkbox', undefined, displayName);
        cb.preferredSize.width = 120;
        cb.value = defaultSettings.lineBreaks && defaultSettings.lineBreaks[field.replace(/ /g, '')] ? 
                   defaultSettings.lineBreaks[field.replace(/ /g, '')].enabled : false;
        
        grp.add('statictext', undefined, 'Replace:');
        var et = grp.add('edittext', undefined, 
                        defaultSettings.lineBreaks && defaultSettings.lineBreaks[field.replace(/ /g, '')] ? 
                        defaultSettings.lineBreaks[field.replace(/ /g, '')].character : '|');
        et.characters = 4;
        grp.add('statictext', undefined, 'with line breaks');
        
        sessionControls[field.replace(/ /g, '')] = { checkbox: cb, charInput: et };
    }
    
    // Chairpersons section
    var pnlChairpersons = tab.add('panel', undefined, 'Chairpersons');
    pnlChairpersons.alignChildren = 'fill';
    pnlChairpersons.margins = [15, 15, 15, 15];
    
    var chairpersonControls = {};
    var chairpersonFields = ['Chairpersons'];
    var chairpersonDisplayNames = ['Chairpersons'];
    
    for (var i = 0; i < chairpersonFields.length; i++) {
        var field = chairpersonFields[i];
        var displayName = chairpersonDisplayNames[i];
        
        var grp = pnlChairpersons.add('group');
        grp.orientation = 'row';
        grp.alignChildren = 'left';
        grp.spacing = 10;
        
        var cb = grp.add('checkbox', undefined, displayName);
        cb.preferredSize.width = 120;
        cb.value = defaultSettings.lineBreaks && defaultSettings.lineBreaks[field.replace(/ /g, '')] ? 
                   defaultSettings.lineBreaks[field.replace(/ /g, '')].enabled : false;
        
        grp.add('statictext', undefined, 'Replace:');
        var et = grp.add('edittext', undefined, 
                        defaultSettings.lineBreaks && defaultSettings.lineBreaks[field.replace(/ /g, '')] ? 
                        defaultSettings.lineBreaks[field.replace(/ /g, '')].character : '|');
        et.characters = 4;
        grp.add('statictext', undefined, 'with line breaks');
        
        chairpersonControls[field.replace(/ /g, '')] = { checkbox: cb, charInput: et };
    }
    
    // Topics section
    var pnlTopics = tab.add('panel', undefined, 'Topics');
    pnlTopics.alignChildren = 'fill';
    pnlTopics.margins = [15, 15, 15, 15];
    
    var topicControls = {};
    var topicFields = ['topicTime', 'topicTitle', 'topicSpeaker'];
    var topicDisplayNames = ['Topic Time', 'Topic Title', 'Topic Speaker'];
    var csvFieldMapping = {
        'topicTime': 'Time',
        'topicTitle': 'Topic Title', 
        'topicSpeaker': 'Speaker'
    };
    
    for (var i = 0; i < topicFields.length; i++) {
        var field = topicFields[i];
        var displayName = topicDisplayNames[i];
        var csvField = csvFieldMapping[field];
        
        var grp = pnlTopics.add('group');
        grp.orientation = 'row';
        grp.alignChildren = 'left';
        grp.spacing = 10;
        
        var cb = grp.add('checkbox', undefined, displayName);
        cb.preferredSize.width = 120;
        cb.value = defaultSettings.lineBreaks && defaultSettings.lineBreaks[field.replace(/ /g, '')] ? 
                   defaultSettings.lineBreaks[field.replace(/ /g, '')].enabled : false;
        
        grp.add('statictext', undefined, 'Replace:');
        var et = grp.add('edittext', undefined, 
                        defaultSettings.lineBreaks && defaultSettings.lineBreaks[field.replace(/ /g, '')] ? 
                        defaultSettings.lineBreaks[field.replace(/ /g, '')].character : '|');
        et.characters = 4;
        grp.add('statictext', undefined, 'with line breaks');
        
        topicControls[field.replace(/ /g, '')] = { checkbox: cb, charInput: et };
    }
    
    // Store references for later access
    tab.sessionControls = sessionControls;
    tab.chairpersonControls = chairpersonControls;
    tab.topicControls = topicControls;
}




// Helper functions for the unified panel
function updateChairpersonsTabFromSettings(tab, settings) {
    if (!settings) return;
    
    if (tab.ddMode) tab.ddMode.selection = settings.mode === 'grid' ? 1 : 0;
    if (tab.rbComma) tab.rbComma.value = settings.inlineSeparator !== 'linebreak';
    if (tab.rbLines) tab.rbLines.value = settings.inlineSeparator === 'linebreak';
    if (tab.ddOrder) tab.ddOrder.selection = settings.order === 'col' ? 1 : 0;
    if (tab.etCols) tab.etCols.text = settings.columns || '2';
    if (tab.etRows) tab.etRows.text = settings.rows || '2';
    if (tab.etColSpace) tab.etColSpace.text = settings.colSpacing || '8';
    if (tab.etRowSpace) tab.etRowSpace.text = settings.rowSpacing || '4';
    if (tab.ddUnits) tab.ddUnits.selection = settings.units ? indexOfExact(['pt', 'mm', 'cm', 'px'], settings.units) : 0;
    if (tab.cbCenter) tab.cbCenter.value = settings.centerGrid || false;
    
    // Update image automation settings
    if (tab.cbEnableImages) tab.cbEnableImages.value = settings.enableImages || false;
    if (tab.etImageFolder) tab.etImageFolder.text = settings.imageFolder || '';
    if (tab.ddImageFitting) tab.ddImageFitting.selection = settings.imageFitting ? indexOfExact(['Use Template Default', 'Fill Frame', 'Fit Proportionally', 'Fit Content to Frame', 'Center Content'], settings.imageFitting) : 0;
    
    // IMPORTANT: Call the UI sync function after updating settings
    // This ensures the image controls are properly enabled/disabled based on current state
    if (typeof tab.syncUI === 'function') {
        tab.syncUI();
    }
}

function updateTopicsTabFromSettings(tab, settings) {
    if (!settings) return;
    
    if (settings.mode === 'table' && tab.rbTable && tab.rbTable.enabled) {
        tab.rbTable.value = true;
    } else if (settings.mode === 'independent' && tab.rbIndependent && tab.rbIndependent.enabled) {
        tab.rbIndependent.value = true;
    }
    if (tab.cbHeader) tab.cbHeader.value = settings.includeHeader !== false;
    if (tab.etSpacing) tab.etSpacing.text = settings.verticalSpacing || '4';
    if (tab.ddSpaceUnits) tab.ddSpaceUnits.selection = settings.units ? indexOfExact(['pt', 'mm', 'cm', 'px'], settings.units) : 0;
}

function updateLineBreaksTabFromSettings(tab, settings) {
    if (!settings || !settings.lineBreaks) return;
    
    // Update session controls
    for (var field in tab.sessionControls) {
        if (settings.lineBreaks[field]) {
            tab.sessionControls[field].checkbox.value = settings.lineBreaks[field].enabled || false;
            tab.sessionControls[field].charInput.text = settings.lineBreaks[field].character || '|';
        }
    }
    
    // Update chairperson controls
    for (var field in tab.chairpersonControls) {
        if (settings.lineBreaks[field]) {
            tab.chairpersonControls[field].checkbox.value = settings.lineBreaks[field].enabled || false;
            tab.chairpersonControls[field].charInput.text = settings.lineBreaks[field].character || '|';
        }
    }
    
    // Update topic controls
    for (var field in tab.topicControls) {
        if (settings.lineBreaks[field]) {
            tab.topicControls[field].checkbox.value = settings.lineBreaks[field].enabled || false;
            tab.topicControls[field].charInput.text = settings.lineBreaks[field].character || '|';
        }
    }
}

// updateImageAutomationTabFromSettings function removed - functionality moved to chairpersons tab

function collectAllSettings(chairTab, topicTab, lineBreakTab) {
    // Collect chairperson settings
    var chairOptions = {
        mode: (chairTab.ddMode.selection.index === 0) ? 'inline' : 'grid',
        inlineSeparator: chairTab.rbLines.value ? 'linebreak' : 'comma',
        order: (chairTab.ddOrder.selection.index === 0) ? 'row' : 'col',
        columns: clampInt(parseInt(chairTab.etCols.text, 10), 1, 99),
        rows: clampInt(parseInt(chairTab.etRows.text, 10), 1, 99),
        colSpacing: chairTab.etColSpace.text,
        rowSpacing: chairTab.etRowSpace.text,
        colSpacingPt: toPointsSafe(chairTab.etColSpace.text, chairTab.ddUnits.selection.text),
        rowSpacingPt: toPointsSafe(chairTab.etRowSpace.text, chairTab.ddUnits.selection.text),
        units: chairTab.ddUnits.selection.text,
        centerGrid: chairTab.cbCenter.value,
        enableImages: chairTab.cbEnableImages.value,
        imageFolder: chairTab.etImageFolder.text,
        imageFitting: chairTab.ddImageFitting.selection ? chairTab.ddImageFitting.selection.text : 'Use Template Default'
    };
    
    // Collect topic settings
    var topicOptions = {
        mode: topicTab.rbTable.value ? 'table' : 'independent',
        includeHeader: topicTab.cbHeader.value,
        verticalSpacing: topicTab.etSpacing.text,
        verticalSpacingPt: toPointsSafe(topicTab.etSpacing.text, topicTab.ddSpaceUnits.selection.text),
        units: topicTab.ddSpaceUnits.selection.text
    };
    
    // Collect line break settings
    var lineBreaks = {};
    
    // Collect session settings
    for (var field in lineBreakTab.sessionControls) {
        lineBreaks[field] = {
            enabled: lineBreakTab.sessionControls[field].checkbox.value,
            character: lineBreakTab.sessionControls[field].charInput.text || '|'
        };
    }
    
    // Collect chairperson settings
    for (var field in lineBreakTab.chairpersonControls) {
        lineBreaks[field] = {
            enabled: lineBreakTab.chairpersonControls[field].checkbox.value,
            character: lineBreakTab.chairpersonControls[field].charInput.text || '|'
        };
    }
    
    // Collect topic settings
    for (var field in lineBreakTab.topicControls) {
        lineBreaks[field] = {
            enabled: lineBreakTab.topicControls[field].checkbox.value,
            character: lineBreakTab.topicControls[field].charInput.text || '|'
        };
    }
    
    var lineBreakOptions = { lineBreaks: lineBreaks };
    
    return {
        chairOptions: chairOptions,
        topicOptions: topicOptions,
        lineBreakOptions: lineBreakOptions
    };
}

function importAllSettings() {
    try {
        var fileExtension = ".json";
        var settingsFile = File.openDialog("Select settings file", "JSON Files:*" + fileExtension);
        
        if (!settingsFile) return null;
        
        settingsFile.encoding = "UTF-8";
        settingsFile.open("r");
        var jsonString = settingsFile.read();
        settingsFile.close();
        
        var settings = jsonFromString(jsonString);
        
        if (!settings) {
            alert("Error: Invalid settings file format.");
            return null;
        }
        
        return settings;
    } catch (e) {
        alert("Error importing settings: " + e);
        return null;
    }
}

function exportAllSettings(allSettings) {
    try {
        var fileExtension = ".json";
        var fileName = "agenda_all_settings" + fileExtension;
        var settingsFile = File.saveDialog("Save all settings as", "JSON Files:*" + fileExtension);
        
        if (!settingsFile) return false;
        
        if (settingsFile.name.indexOf(fileExtension) === -1) {
            settingsFile = new File(settingsFile.absoluteURI + fileExtension);
        }
        
        var jsonString = settingsToJSON(allSettings);
        
        settingsFile.encoding = "UTF-8";
        settingsFile.open("w");
        settingsFile.write(jsonString);
        settingsFile.close();
        
        return true;
    } catch (e) {
        alert("Error exporting settings: " + e);
        return false;
    }
}

/* ---------------- UTILITY FUNCTIONS ---------------- */

// Old getLineBreakOptions function removed - replaced by unified settings panel

// Old getChairLayoutOptions function removed - replaced by unified settings panel

// Old getTopicLayoutOptions function removed - replaced by unified settings panel

function clampInt(v, min, max) { if (isNaN(v)) v = min; if (v < min) v = min; if (v > max) v = max; return v; }

function toPointsSafe(txt, units) { 
    var n = parseFloat(txt); 
    if (isNaN(n)) n = 0; 
    switch(units) {
        case 'mm': return n * 2.834645669;
        case 'cm': return n * 28.34645669;
        case 'px': return n * 0.75; // 96 DPI to 72 DPI conversion
        case 'pt':
        default: return n;
    }
}

/* ---------------- PAGE & PLACEHOLDER FILLING ---------------- */

function fillPagePlaceholders(page, session, chairOptions, topicOptions, lineBreakOptions, availableFields) {
    var hasIndependentTopicLayout = detectIndependentTopicSetup(page);
    var hasTableTopicLayout = !!findItemByLabel(page, "topicsTable");
    var placeholders = findPlaceholdersManually(page, availableFields);

    if (!placeholders) {
        alert("Error: Could not find any placeholders on the duplicated page.");
        return;
    }

    // Fill session-level placeholders with line break support
    if (placeholders.sessionTitle && session.title) {
        placeholders.sessionTitle.contents = session.title;
        if (lineBreakOptions.lineBreaks.SessionTitle && lineBreakOptions.lineBreaks.SessionTitle.enabled) {
            applyLineBreaksToFrame(placeholders.sessionTitle, lineBreakOptions.lineBreaks.SessionTitle.character);
        }
    }
    if (placeholders.sessionTime && session.time) {
        placeholders.sessionTime.contents = session.time;
        if (lineBreakOptions.lineBreaks.SessionTime && lineBreakOptions.lineBreaks.SessionTime.enabled) {
            applyLineBreaksToFrame(placeholders.sessionTime, lineBreakOptions.lineBreaks.SessionTime.character);
        }
    }
    if (placeholders.sessionNo && session.no) {
        placeholders.sessionNo.contents = session.no;
        if (lineBreakOptions.lineBreaks.SessionNo && lineBreakOptions.lineBreaks.SessionNo.enabled) {
            applyLineBreaksToFrame(placeholders.sessionNo, lineBreakOptions.lineBreaks.SessionNo.character);
        }
    }
    
    // Fill chairpersons
    if (placeholders.chairpersons && session.chairs) {
        populateChairs(page, placeholders.chairpersons, session.chairs, chairOptions, lineBreakOptions);
    }

    // Fill topics
    if (topicOptions.mode === 'table' && placeholders.topicsTable) {
        populateTopicsAsTable(placeholders.topicsTable, session.topics, topicOptions, lineBreakOptions);
    } else if (topicOptions.mode === 'independent' && hasIndependentTopicLayout) {
        populateTopicsAsIndependentRows(page, session.topics, topicOptions, lineBreakOptions);
    }
}

function populateChairs(page, chairFrame, chairsRaw, opts, lineBreakOptions) {
    removeExistingChairClones(page);
    var names = splitChairs(chairsRaw); // Uses || delimiter
    
    // Initialize image results if image automation is enabled
    if (opts.enableImages && opts.imageFolder) {
        imageResults.enabled = true;
        imageResults.folder = opts.imageFolder;
        // Only set total if not already set (avoid overwriting from previous sessions)
        if (imageResults.totalAttempted === 0) {
            imageResults.totalAttempted = names.length;
        } else {
            imageResults.totalAttempted += names.length;
        }
    }
    
    if (opts.mode === 'inline') {
        var content = (names.length === 0) ? "" : ((opts.inlineSeparator === 'linebreak') ? names.join("\r") : names.join(", "));
        chairFrame.contents = content;
        
        // Apply line breaks if requested
        if (lineBreakOptions.lineBreaks.Chairpersons && lineBreakOptions.lineBreaks.Chairpersons.enabled && content) {
            applyLineBreaksToFrame(chairFrame, lineBreakOptions.lineBreaks.Chairpersons.character);
        }
        return;
    }
    
    // For grid mode, handle the original prototype
    var protoRoot = getPrototypeRoot(chairFrame);
    var rootIsGroup = (protoRoot.constructor.name === "Group");
    var gb = protoRoot.geometricBounds;
    var w = gb[3] - gb[1], h = gb[2] - gb[0];
    var total = names.length;
    
    // Hide/clear the original prototype
    if (rootIsGroup) {
        protoRoot.visible = false;
    } else {
        chairFrame.contents = "";
    }
    
    if (total === 0) return;
    
    var cols, rows;
    if (opts.order === 'row') { 
        cols = Math.max(1, opts.columns); 
        rows = Math.ceil(total / cols); 
    } else { 
        rows = Math.max(1, opts.rows); 
        cols = Math.ceil(total / rows); 
    }
    
    var startX, startY;
    if (opts.centerGrid) {
        // Center based on ACTUAL chairpersons, not full grid
        var pageBounds = page.bounds; // [top, left, bottom, right]
        var pageWidth = pageBounds[3] - pageBounds[1];
        
        // Calculate how many chairpersons will be in each row
        var actualColsInLastRow = total % cols;
        if (actualColsInLastRow === 0) actualColsInLastRow = cols;
        
        // Calculate the width of the actual chairpersons that will be placed
        var actualGridWidth;
        if (opts.order === 'row') {
            // For row-first: calculate based on the widest row (usually the first row)
            var firstRowCols = Math.min(cols, total);
            actualGridWidth = (firstRowCols - 1) * opts.colSpacingPt + firstRowCols * w;
        } else {
            // For column-first: calculate based on actual columns used
            var actualCols = Math.min(cols, total);
            actualGridWidth = (actualCols - 1) * opts.colSpacingPt + actualCols * w;
        }
        
        // Center the actual chairpersons horizontally
        startX = pageBounds[1] + (pageWidth - actualGridWidth) / 2;
        startY = gb[0]; // Keep original vertical position
    } else {
        // Use the current chairperson frame's position as the starting point
        startX = gb[1];
        startY = gb[0];
    }
    
    for (var i = 0; i < total; i++) {
        var r, c;
        if (opts.order === 'row') { 
            r = Math.floor(i / cols); 
            c = i % cols; 
        } else { 
            c = Math.floor(i / rows); 
            r = i % rows; 
        }
        
        var targetX = startX + c * (w + opts.colSpacingPt);
        var targetY = startY + r * (h + opts.rowSpacingPt);
        
        var clone = protoRoot.duplicate(page);
        clone.label = "chairClone";
        clone.visible = true; // Ensure clone is visible
        
        var tf = rootIsGroup ? findDescendantByLabel(clone, "chairpersons") : clone;
        if (tf) {
            tf.contents = names[i];
            // Apply line breaks if requested
            if (lineBreakOptions.lineBreaks.Chairpersons && lineBreakOptions.lineBreaks.Chairpersons.enabled) {
                applyLineBreaksToFrame(tf, lineBreakOptions.lineBreaks.Chairpersons.character);
            }
        }
        
        // Process images if automation is enabled and we have a group
        if (opts.enableImages && opts.imageFolder && rootIsGroup) {
            var imageResult = processChairpersonImages(clone, names[i], opts.imageFolder, opts.imageFitting);
            
            // Track results for reporting
            if (imageResult.avatar || imageResult.flag) {
                var successMsg = imageResult.cleanName + ' -> ';
                if (imageResult.avatar) successMsg += imageResult.avatarFile;
                if (imageResult.flag) successMsg += ' + ' + imageResult.flagFile;
                imageResults.successful.push(successMsg);
            } else {
                imageResults.failed.push({
                    name: imageResult.cleanName || names[i],
                    attempts: imageResult.attempts
                });
            }
        }
        
        var cgb = clone.geometricBounds;
        clone.move(undefined, [targetX - cgb[1], targetY - cgb[0]]);
    }
}

function populateTopicsAsTable(tableFrame, topics, opts, lineBreakOptions) {
    try { while (tableFrame.tables.length) tableFrame.tables[0].remove(); } catch (e) { }
    tableFrame.contents = "";
    if (topics.length === 0) return;
    var headerRows = opts.includeHeader ? 1 : 0;
    var tbl = tableFrame.insertionPoints[0].tables.add({ headerRowCount: headerRows, bodyRowCount: topics.length, columnCount: 3 });
    var frameWidth = tableFrame.geometricBounds[3] - tableFrame.geometricBounds[1];
    tbl.columns[0].width = frameWidth * 0.20;
    tbl.columns[1].width = frameWidth * 0.50;
    tbl.columns[2].width = frameWidth * 0.30;
    if (opts.includeHeader) {
        tbl.rows[0].cells[0].contents = "Time";
        tbl.rows[0].cells[1].contents = "Topic";
        tbl.rows[0].cells[2].contents = "Speaker";
    }
    for (var i = 0; i < topics.length; i++) {
        var rIndex = headerRows ? i + 1 : i;
        tbl.rows[rIndex].cells[0].contents = topics[i].time;
        if (lineBreakOptions.lineBreaks.topicTime && lineBreakOptions.lineBreaks.topicTime.enabled) {
            applyLineBreaksToCell(tbl.rows[rIndex].cells[0], lineBreakOptions.lineBreaks.topicTime.character);
        }
        
        tbl.rows[rIndex].cells[1].contents = topics[i].title;
        if (lineBreakOptions.lineBreaks.topicTitle && lineBreakOptions.lineBreaks.topicTitle.enabled) {
            applyLineBreaksToCell(tbl.rows[rIndex].cells[1], lineBreakOptions.lineBreaks.topicTitle.character);
        }
        
        tbl.rows[rIndex].cells[2].contents = topics[i].speaker;
        if (lineBreakOptions.lineBreaks.topicSpeaker && lineBreakOptions.lineBreaks.topicSpeaker.enabled) {
            applyLineBreaksToCell(tbl.rows[rIndex].cells[2], lineBreakOptions.lineBreaks.topicSpeaker.character);
        }
    }
}

function populateTopicsAsIndependentRows(page, topics, opts, lineBreakOptions) {
    removeExistingTopicClones(page);
    
    // Find the topic frames with new naming convention
    var protoTime = findItemByLabel(page, "topicTime");
    var protoTopic = findItemByLabel(page, "topicTitle");
    var protoSpeaker = findItemByLabel(page, "topicSpeaker");
    
    if (!protoTime || !protoTopic || !protoSpeaker) {
        alert("Template Error: Could not find frames labeled 'topicTime', 'topicTitle', and 'topicSpeaker' for independent topic layout.");
        return;
    }

    // Check if these frames are all in the same group
    var parentGroup = findCommonParentGroup([protoTime, protoTopic, protoSpeaker]);
    
    if (parentGroup) {
        // Group-based approach: duplicate the entire group
        var gb = parentGroup.geometricBounds;
        var groupHeight = gb[2] - gb[0];
        var baseX = gb[1];
        var baseY = gb[0];
        
        // Hide the original prototype group
        parentGroup.visible = false;

        for (var i = 0; i < topics.length; i++) {
            var cloneGroup = parentGroup.duplicate(page);
            cloneGroup.label = "topicCloneGroup";
            cloneGroup.visible = true;
            
            // Find and populate the topic elements within the cloned group
            var tTime = findDescendantByLabel(cloneGroup, "topicTime");
            var tTopic = findDescendantByLabel(cloneGroup, "topicTitle");
            var tSpeaker = findDescendantByLabel(cloneGroup, "topicSpeaker");
            
            if (tTime) {
                tTime.contents = topics[i].time;
                if (lineBreakOptions.lineBreaks.topicTime && lineBreakOptions.lineBreaks.topicTime.enabled) {
                    applyLineBreaksToFrame(tTime, lineBreakOptions.lineBreaks.topicTime.character);
                }
            }
            if (tTopic) {
                tTopic.contents = topics[i].title;
                if (lineBreakOptions.lineBreaks.topicTitle && lineBreakOptions.lineBreaks.topicTitle.enabled) {
                    applyLineBreaksToFrame(tTopic, lineBreakOptions.lineBreaks.topicTitle.character);
                }
            }
            if (tSpeaker) {
                tSpeaker.contents = topics[i].speaker;
                if (lineBreakOptions.lineBreaks.topicSpeaker && lineBreakOptions.lineBreaks.topicSpeaker.enabled) {
                    applyLineBreaksToFrame(tSpeaker, lineBreakOptions.lineBreaks.topicSpeaker.character);
                }
            }
            
            // Position the entire group
            var targetY = baseY + i * (groupHeight + opts.verticalSpacingPt);
            var currentBounds = cloneGroup.geometricBounds;
            cloneGroup.move(undefined, [baseX - currentBounds[1], targetY - currentBounds[0]]);
        }
    } else {
        // Individual frame approach: frames are not grouped
        var baseY = protoTime.geometricBounds[0];
        var rowHeight = protoTime.geometricBounds[2] - protoTime.geometricBounds[0];
        
        // Hide original prototypes
        protoTime.visible = false;
        protoTopic.visible = false;
        protoSpeaker.visible = false;

        for (var i = 0; i < topics.length; i++) {
            var targetY = baseY + i * (rowHeight + opts.verticalSpacingPt);
            
            // Clone and position Time frame
            var timeClone = protoTime.duplicate(page);
            timeClone.label = "topicClone_Time";
            timeClone.visible = true;
            timeClone.contents = topics[i].time;
            if (lineBreakOptions.lineBreaks.topicTime && lineBreakOptions.lineBreaks.topicTime.enabled) {
                applyLineBreaksToFrame(timeClone, lineBreakOptions.lineBreaks.topicTime.character);
            }
            timeClone.move(undefined, [0, targetY - timeClone.geometricBounds[0]]);
            
            // Clone and position Topic frame
            var topicClone = protoTopic.duplicate(page);
            topicClone.label = "topicClone_Topic";
            topicClone.visible = true;
            topicClone.contents = topics[i].title;
            if (lineBreakOptions.lineBreaks.topicTitle && lineBreakOptions.lineBreaks.topicTitle.enabled) {
                applyLineBreaksToFrame(topicClone, lineBreakOptions.lineBreaks.topicTitle.character);
            }
            topicClone.move(undefined, [0, targetY - topicClone.geometricBounds[0]]);
            
            // Clone and position Speaker frame
            var speakerClone = protoSpeaker.duplicate(page);
            speakerClone.label = "topicClone_Speaker";
            speakerClone.visible = true;
            speakerClone.contents = topics[i].speaker;
            if (lineBreakOptions.lineBreaks.topicSpeaker && lineBreakOptions.lineBreaks.topicSpeaker.enabled) {
                applyLineBreaksToFrame(speakerClone, lineBreakOptions.lineBreaks.topicSpeaker.character);
            }
            speakerClone.move(undefined, [0, targetY - speakerClone.geometricBounds[0]]);
        }
    }
}

/* ---------------- LINE BREAK FUNCTIONS ---------------- */

function applyLineBreaksToFrame(textFrame, character) {
    try {
        if (!textFrame || !textFrame.hasOwnProperty("contents")) return;
        var content = textFrame.contents;
        if (!content || content === "") return;
        
        // Replace the specified character with line breaks
        var newContent = content.replace(new RegExp("\\"+character, "g"), "\r");
        textFrame.contents = newContent;
    } catch (e) {
        // Silently handle any errors to prevent script failure
    }
}

function applyLineBreaksToCell(cell, character) {
    try {
        if (!cell || !cell.hasOwnProperty("contents")) return;
        var content = cell.contents;
        if (!content || content === "") return;
        
        // Replace the specified character with line breaks
        var newContent = content.replace(new RegExp("\\"+character, "g"), "\r");
        cell.contents = newContent;
    } catch (e) {
        // Silently handle any errors to prevent script failure
    }
}

/* ---------------- IMAGE AUTOMATION FUNCTIONS ---------------- */

// Global variable to track image placement results for reporting
var imageResults = {
    enabled: false,
    folder: '',
    successful: [],
    failed: [],
    totalAttempted: 0
};

// Extract clean name from CSV format "Prof. Dr. John Smith|CEO" -> "John Smith"
function extractCleanName(rawName) {
    if (!rawName) return '';
    
    // Remove title part: "Prof. Dr. John Smith|CEO" -> "Prof. Dr. John Smith"
    var nameOnly = rawName.split('|')[0];
    
    // Strip academic and professional titles
    return stripTitles(trimES(nameOnly));
}

// Strip academic titles, professional titles, and medical suffixes
function stripTitles(name) {
    if (!name) return '';
    
    // Define patterns to remove (case insensitive)
    var prefixPatterns = [
        /^Prof\.?\s+Dr\.?\s+Med\.?\s+/i,  // Prof. Dr. Med.
        /^Prof\.?\s+Dr\.?\s+/i,          // Prof. Dr.
        /^Assoc\.?\s+Prof\.?\s+/i,       // Assoc. Prof. / Associate Prof.
        /^Associate\s+Prof\.?\s+/i,      // Associate Prof.
        /^Prof\.?\s+/i,                  // Prof.
        /^Dr\.?\s+/i,                    // Dr.
        /^Mr\.?\s+/i,                    // Mr.
        /^Mrs\.?\s+/i,                   // Mrs.
        /^Ms\.?\s+/i,                    // Ms.
        /^Miss\.?\s+/i                   // Miss
    ];
    
    var suffixPatterns = [
        /\s+MD\.?$/i,                    // MD at end
        /\s+Ph\.?D\.?$/i,                // PhD at end
        /\s+M\.?D\.?$/i                  // M.D. at end
    ];
    
    var cleanName = name;
    
    // Remove prefixes
    for (var i = 0; i < prefixPatterns.length; i++) {
        cleanName = cleanName.replace(prefixPatterns[i], '');
    }
    
    // Remove suffixes
    for (var j = 0; j < suffixPatterns.length; j++) {
        cleanName = cleanName.replace(suffixPatterns[j], '');
    }
    
    return trimES(cleanName);
}

// Generate name variations for image matching
function generateNameVariations(cleanName) {
    if (!cleanName) return [];
    
    var variations = [];
    var baseName = cleanName.toLowerCase();
    
    // Add exact case variations
    variations.push(cleanName);           // Original case
    variations.push(baseName);            // Lowercase
    variations.push(cleanName.toUpperCase()); // Uppercase
    
    // Add spacing variations
    var noSpaces = baseName.replace(/\s+/g, '');
    var underscores = baseName.replace(/\s+/g, '_');
    var hyphens = baseName.replace(/\s+/g, '-');
    
    variations.push(noSpaces);
    variations.push(underscores);
    variations.push(hyphens);
    
    // Add title case variation
    var titleCase = cleanName.replace(/\w\S*/g, function(txt) {
        return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    });
    variations.push(titleCase);
    variations.push(titleCase.replace(/\s+/g, '_'));
    variations.push(titleCase.replace(/\s+/g, '-'));
    
    // Remove duplicates
    var unique = [];
    for (var i = 0; i < variations.length; i++) {
        var found = false;
        for (var j = 0; j < unique.length; j++) {
            if (unique[j] === variations[i]) {
                found = true;
                break;
            }
        }
        if (!found) {
            unique.push(variations[i]);
        }
    }
    
    return unique;
}

// Generate flag name variations based on avatar name
function generateFlagVariations(cleanName) {
    if (!cleanName) return [];
    
    var nameVariations = generateNameVariations(cleanName);
    var flagVariations = [];
    
    for (var i = 0; i < nameVariations.length; i++) {
        var name = nameVariations[i];
        
        // Hyphen variations
        flagVariations.push('flag-' + name);
        flagVariations.push(name + '-flag');
        
        // Underscore variations
        flagVariations.push('flag_' + name);
        flagVariations.push(name + '_flag');
        
        // No separator variations
        flagVariations.push('flag' + name);
        flagVariations.push(name + 'flag');
        
        // Space variations (new enhancement)
        flagVariations.push('flag ' + name);
        flagVariations.push(name + ' flag');
        
        // Additional space variations for common patterns
        flagVariations.push('Flag ' + name);
        flagVariations.push(name + ' Flag');
        flagVariations.push('FLAG ' + name);
        flagVariations.push(name + ' FLAG');
    }
    
    return flagVariations;
}

// Find matching image file in folder
function findMatchingImageFile(nameVariations, imageFolder, imageType) {
    if (!imageFolder || !nameVariations || nameVariations.length === 0) {
        return { found: false, filePath: '', searchAttempts: [] };
    }
    
    var folder = new Folder(imageFolder);
    if (!folder.exists) {
        return { found: false, filePath: '', searchAttempts: ['Folder does not exist: ' + imageFolder] };
    }
    
    var supportedFormats = ['.jpg', '.jpeg', '.png', '.tif', '.tiff', '.psd'];
    var searchAttempts = [];
    
    // Try each name variation with each format
    for (var i = 0; i < nameVariations.length; i++) {
        var nameVar = nameVariations[i];
        for (var j = 0; j < supportedFormats.length; j++) {
            var format = supportedFormats[j];
            var fileName = nameVar + format;
            var filePath = imageFolder + '/' + fileName;
            
            searchAttempts.push(fileName);
            
            var file = new File(filePath);
            if (file.exists) {
                return {
                    found: true,
                    filePath: file.fsName,
                    fileName: fileName,
                    searchAttempts: searchAttempts
                };
            }
        }
    }
    
    return {
        found: false,
        filePath: '',
        searchAttempts: searchAttempts
    };
}

// Place image into labeled frame within a group
function placeImageInFrame(groupItem, frameLabel, imagePath, fittingOption) {
    try {
        var imageFrame = findDescendantByLabel(groupItem, frameLabel);
        if (!imageFrame) {
            return false; // Frame not found
        }
        
        // Place the image
        var imageFile = new File(imagePath);
        if (!imageFile.exists) {
            return false;
        }
        
        imageFrame.place(imageFile);
        
        // Apply fitting based on user selection
        try {
            if (fittingOption && fittingOption !== 'Use Template Default') {
                // First, get the placed graphic
                if (imageFrame.graphics.length > 0) {
                    var graphic = imageFrame.graphics[0];
                    
                    switch (fittingOption) {
                        case 'Fill Frame':
                            graphic.fit(FitOptions.FILL_PROPORTIONALLY);
                            break;
                        case 'Fit Proportionally':
                            graphic.fit(FitOptions.FIT_CONTENT_PROPORTIONALLY);
                            break;
                        case 'Fit Content to Frame':
                            graphic.fit(FitOptions.CONTENT_TO_FRAME);
                            break;
                        case 'Center Content':
                            graphic.fit(FitOptions.CENTER_CONTENT);
                            break;
                    }
                }
            }
            // If 'Use Template Default' or no option, don't change existing fitting
        } catch (fitError) {
            // Fitting might fail on some frame types, but image is still placed
            $.writeln("Fitting error: " + fitError);
        }
        
        return true;
    } catch (e) {
        return false;
    }
}

// Process images for a single chairperson
function processChairpersonImages(groupItem, rawName, imageFolder, fittingOption) {
    if (!imageFolder || !rawName) {
        return { avatar: false, flag: false, attempts: [] };
    }
    
    var cleanName = extractCleanName(rawName);
    if (!cleanName) {
        return { avatar: false, flag: false, attempts: ['Could not extract clean name from: ' + rawName] };
    }
    
    var results = { avatar: false, flag: false, attempts: [], cleanName: cleanName };
    
    // Try to find and place avatar image
    var avatarVariations = generateNameVariations(cleanName);
    var avatarResult = findMatchingImageFile(avatarVariations, imageFolder, 'avatar');
    
    results.attempts = results.attempts.concat(avatarResult.searchAttempts);
    
    if (avatarResult.found) {
        var avatarPlaced = placeImageInFrame(groupItem, 'chairAvatar', avatarResult.filePath, fittingOption);
        if (avatarPlaced) {
            results.avatar = true;
            results.avatarFile = avatarResult.fileName;
            
            // If avatar was successful, try to find and place flag
            var flagVariations = generateFlagVariations(cleanName);
            var flagResult = findMatchingImageFile(flagVariations, imageFolder, 'flag');
            
            if (flagResult.found) {
                var flagPlaced = placeImageInFrame(groupItem, 'chairFlag', flagResult.filePath, fittingOption);
                if (flagPlaced) {
                    results.flag = true;
                    results.flagFile = flagResult.fileName;
                }
            }
        }
    }
    
    return results;
}

/* ---------------- HELPERS & UTILITIES ---------------- */

function removeExistingChairClones(page) {
    var items = page.allPageItems;
    for (var i = items.length - 1; i >= 0; i--) { 
        try { 
            if (items[i].label === "chairClone") items[i].remove(); 
        } catch (e) { } 
    }
}

function removeExistingTopicClones(page) {
    var items = page.allPageItems;
    for (var i = items.length - 1; i >= 0; i--) { 
        try { 
            if (items[i].label && (items[i].label === "topicCloneGroup" || items[i].label.indexOf("topicClone_") === 0)) {
                items[i].remove(); 
            }
        } catch (e) { } 
    }
}

function getPrototypeRoot(chairFrame) {
    try { if (chairFrame.parent.constructor.name === "Group") return chairFrame.parent; } catch (e) { }
    return chairFrame;
}

function splitChairs(raw) {
    var out = []; 
    if (!raw) return out;
    // Changed from | to || as delimiter
    var parts = raw.split("||");
    for (var i = 0; i < parts.length; i++) { 
        var s = trimES(parts[i]); 
        if (s !== "") out.push(s); 
    }
    return out;
}

function findItemByLabel(page, label) {
    var items = page.allPageItems;
    for (var i = 0; i < items.length; i++) { 
        try {
            if (items[i].label === label) return items[i]; 
        } catch (e) { } 
    }
    return null;
}

function findDescendantByLabel(rootItem, label) {
    if (!rootItem.allPageItems) return null;
    var arr = rootItem.allPageItems;
    for (var i = 0; i < arr.length; i++) { 
        try { 
            if (arr[i].label === label) return arr[i]; 
        } catch (e) { } 
    }
    return null;
}

function findCommonParentGroup(frames) {
    if (!frames || frames.length === 0) return null;
    
    // Check if all frames have the same parent and that parent is a Group
    var firstParent = null;
    try {
        if (frames[0].parent && frames[0].parent.constructor.name === "Group") {
            firstParent = frames[0].parent;
        } else {
            return null; // First frame is not in a group
        }
    } catch (e) {
        return null;
    }
    
    // Check if all other frames have the same parent
    for (var i = 1; i < frames.length; i++) {
        try {
            if (!frames[i].parent || frames[i].parent !== firstParent) {
                return null; // Different parent or not in a group
            }
        } catch (e) {
            return null;
        }
    }
    
    return firstParent;
}

function findPlaceholdersManually(page, availableFields) {
    var placeholders = {
        sessionTitle: null,
        sessionTime: null,
        sessionNo: null,
        chairpersons: null,
        topicsTable: null
    };
    
    var all = page.allPageItems;
    for (var i = 0; i < all.length; i++) {
        var it = all[i];
        var lbl = ""; 
        try { lbl = it.label; } catch (e) { }
        if (!lbl) continue;
        
        // Check for session-level placeholders
        if (lbl === "sessionTitle" && it.hasOwnProperty("parentStory")) {
            placeholders.sessionTitle = it;
        } else if (lbl === "sessionTime" && it.hasOwnProperty("parentStory")) {
            placeholders.sessionTime = it;
        } else if (lbl === "sessionNo" && it.hasOwnProperty("parentStory")) {
            placeholders.sessionNo = it;
        } else if (lbl === "chairpersons" && it.hasOwnProperty("parentStory")) {
            placeholders.chairpersons = it;
        } else if (lbl === "topicsTable" && it.hasOwnProperty("parentStory")) {
            placeholders.topicsTable = it;
        }
    }
    
    // Check if we have at least one required placeholder
    var hasAnyPlaceholder = false;
    for (var key in placeholders) {
        if (placeholders[key]) {
            hasAnyPlaceholder = true;
            break;
        }
    }
    
    if (!hasAnyPlaceholder) {
        alert("Template Error:\nCould not find any labeled placeholders on Page 1.\n\nAvailable labels should match CSV columns:\nsessionTitle, sessionTime, sessionNo, chairpersons, topicsTable");
        return null;
    }

    return placeholders;
}

/* ---------------- SETTINGS MANAGEMENT ---------------- */

function getDefaultSettings() {
    // Default settings
    return {
        chairMode: 'inline',
        chairInlineSeparator: 'comma',
        chairOrder: 'row',
        chairColumns: 2,
        chairRows: 2,
        chairColSpacing: 8,
        chairRowSpacing: 4,
        chairUnits: 'pt',
        chairCenterGrid: false,
        chairApplyLineBreaks: true,
        chairLineBreakChar: '|',
        topicMode: null,
        topicIncludeHeader: true,
        topicVerticalSpacing: 4,
        topicUnits: 'pt',
        topicApplyLineBreaks: true,
        topicLineBreakChar: '|',
        lineBreaks: {
            SessionTitle: { enabled: false, character: '|' },
            SessionTime: { enabled: false, character: '|' },
            SessionNo: { enabled: false, character: '|' },
            Chairpersons: { enabled: false, character: '|' },
            topicTime: { enabled: false, character: '|' },
            topicTitle: { enabled: false, character: '|' },
            topicSpeaker: { enabled: false, character: '|' }
        }
    };
}


// Old exportSettings function removed - replaced by centralized function in unified panel


// Old importSettings function removed - replaced by centralized function in unified panel

/**
 * Convert settings object to JSON string
 * @param {Object} settings - The settings object
 * @return {String} - JSON string representation
 */
function settingsToJSON(settings) {
    // Complete JSON stringifier for ExtendScript
    function stringifyValue(value) {
        if (value === null) {
            return "null";
        } else if (typeof value === "string") {
            return "\"" + value.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n').replace(/\r/g, '\\r').replace(/\t/g, '\\t') + "\"";
        } else if (typeof value === "number") {
            return value.toString();
        } else if (typeof value === "boolean") {
            return value ? "true" : "false";
        } else if (value && typeof value === "object" && value.constructor === Array) {
            var arrayStr = "[";
            for (var i = 0; i < value.length; i++) {
                if (i > 0) arrayStr += ",";
                arrayStr += stringifyValue(value[i]);
            }
            arrayStr += "]";
            return arrayStr;
        } else if (typeof value === "object") {
            var objStr = "{";
            var first = true;
            for (var key in value) {
                if (value.hasOwnProperty(key)) {
                    if (!first) objStr += ",";
                    first = false;
                    objStr += "\"" + key + "\":" + stringifyValue(value[key]);
                }
            }
            objStr += "}";
            return objStr;
        } else {
            return "null";
        }
    }
    
    return stringifyValue(settings);
}

/**
 * Parse JSON string to object
 * @param {String} jsonString - The JSON string to parse
 * @return {Object|null} - The parsed object or null if parsing failed
 */
function jsonFromString(jsonString) {
    try {
        // For ExtendScript compatibility, use eval with safety checks
        if (typeof jsonString !== "string" || !jsonString) return null;
        
        // Basic validation to ensure it's a JSON object
        jsonString = jsonString.replace(/^\s+|\s+$/g, ""); // trim
        if (jsonString.charAt(0) !== "{" || jsonString.charAt(jsonString.length - 1) !== "}") {
            return null;
        }
        
        // Additional safety checks for common JSON patterns
        if (jsonString.indexOf("function") !== -1 || jsonString.indexOf("eval") !== -1) {
            return null; // Prevent execution of potentially dangerous code
        }
        
        // Use eval in a controlled way
        var result = eval("(" + jsonString + ")");
        
        // Validate the result is an object and not a function
        if (typeof result === "object" && result !== null && typeof result !== "function") {
            return result;
        } else {
            return null;
        }
    } catch (e) {
        $.writeln("JSON parsing error: " + e);
        return null;
    }
}

/* ---------------- EXPORT REPORT ---------------- */

function exportReport(sessionsData) {
    // Get a file path from the user with a default name
    var reportFilePath = File.saveDialog("Save Report As", "Text Files:*.txt");
    if (!reportFilePath) {
        alert("Report generation cancelled.");
        return null;
    }
    
    // Ensure the file has a .txt extension
    if (!reportFilePath.toString().match(/\.txt$/i)) {
        reportFilePath = new File(reportFilePath.toString() + ".txt");
    } else {
        reportFilePath = new File(reportFilePath);
    }
    
    // Create and open the file
    var reportFile = reportFilePath;
    reportFile.encoding = "UTF-8";
    
    // Make sure we can write to the file
    if (!reportFile.open("w")) {
        alert("Could not create report file. Please check file permissions.");
        return null;
    }
    
    // Write the report header
    reportFile.writeln("AGENDA IMPORT REPORT");
    reportFile.writeln("====================");
    reportFile.writeln("Date: " + new Date().toLocaleString());
    reportFile.writeln("Sessions imported: " + sessionsData.length);
    reportFile.writeln("");
    
    // Check for overset text errors
    var doc = app.activeDocument;
    var oversetErrors = [];
    
    // Check all text frames in the document for overset text
    for (var p = 0; p < doc.pages.length; p++) {
        var page = doc.pages[p];
        
        // First, check direct text frames on the page
        for (var t = 0; t < page.textFrames.length; t++) {
            var textFrame = page.textFrames[t];
            try {
                if (textFrame.overflows) {
                    var frameLabel = textFrame.label || "Unlabeled";
                    oversetErrors.push({
                        page: p + 1,
                        frame: frameLabel,
                        location: "Page " + (p + 1) + ", " + frameLabel + " text frame"
                    });
                }
            } catch (e) {
                $.writeln("Error checking direct frame: " + e);
            }
        }
        
        // Then check text frames in groups
        var groups = page.groups;
        for (var g = 0; g < groups.length; g++) {
            var group = groups[g];
            for (var gt = 0; gt < group.textFrames.length; gt++) {
                var groupTextFrame = group.textFrames[gt];
                try {
                    if (groupTextFrame.overflows) {
                        var groupFrameLabel = groupTextFrame.label || "Unlabeled (in group)";
                        oversetErrors.push({
                            page: p + 1,
                            frame: groupFrameLabel,
                            location: "Page " + (p + 1) + ", " + groupFrameLabel + " text frame"
                        });
                    }
                } catch (e) {
                    $.writeln("Error checking group frame: " + e);
                }
            }
        }
        
        // Force check for any text frame with overset text indicator
    try {
        var pageItems = page.pageItems;
        for (var pi = 0; pi < pageItems.length; pi++) {
            var item = pageItems[pi];
            if (item.constructor.name === "TextFrame") {
                // Check using the overflows property
                if (item.overflows) {
                    var itemLabel = item.label || "Unlabeled (page item)";
                    oversetErrors.push({
                        page: p + 1,
                        frame: itemLabel,
                        location: "Page " + (p + 1) + ", " + itemLabel + " text frame"
                    });
                }
                
                // Alternative check: see if there's an overset point
                try {
                    if (item.contents && item.overflowText && item.overflowText.length > 0) {
                        var altLabel = item.label || "Unlabeled (overflow text)";
                        var alreadyReported = false;
                        
                        // Check if we already reported this frame
                        for (var oe = 0; oe < oversetErrors.length; oe++) {
                            if (oversetErrors[oe].frame === altLabel && 
                                oversetErrors[oe].page === p + 1) {
                                alreadyReported = true;
                                break;
                            }
                        }
                        
                        if (!alreadyReported) {
                            oversetErrors.push({
                                page: p + 1,
                                frame: altLabel,
                                location: "Page " + (p + 1) + ", " + altLabel + " text frame (overflow text)"
                            });
                        }
                    }
                } catch (innerE) {
                    // Ignore errors in alternative check
                }
            }
        }
    } catch (e) {
        $.writeln("Error in page items check: " + e);
    }
    }
    
    // Create a test text frame with overset text for demonstration if requested
    var createTestOverset = false; // Set to true to create a test frame with overset text
    if (createTestOverset) {
        try {
            var testFrame = doc.pages[0].textFrames.add({geometricBounds: [20, 20, 30, 30]});
            testFrame.label = "TEST_OVERSET";
            testFrame.contents = "This is a test text frame with deliberately overset text that should be detected by the script. This text is intentionally long to ensure it overflows the small frame we created.";
            $.writeln("Created test overset text frame");
        } catch (e) {
            $.writeln("Error creating test frame: " + e);
        }
    }
    
    // Always report overset text check results - put at the beginning for visibility
    reportFile.writeln("\n\n*** OVERSET TEXT CHECK ***");
    reportFile.writeln("=========================");
    
    // Debug info about the document
    reportFile.writeln("Document has " + doc.pages.length + " pages");
    reportFile.writeln("Checked all text frames for overflow");
    
    // Force a final check for any overset text
    var finalCheck = [];
    app.doScript(function() {
        for (var p = 0; p < doc.pages.length; p++) {
            var page = doc.pages[p];
            var frames = page.textFrames;
            for (var f = 0; f < frames.length; f++) {
                if (frames[f].overflows) {
                    finalCheck.push("Page " + (p+1) + ", " + (frames[f].label || "Unlabeled") + " text frame");
                }
            }
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.FAST_ENTIRE_SCRIPT);
    
    // Add any newly found errors
    for (var fc = 0; fc < finalCheck.length; fc++) {
        var found = false;
        for (var oe = 0; oe < oversetErrors.length; oe++) {
            if (oversetErrors[oe].location === finalCheck[fc]) {
                found = true;
                break;
            }
        }
        if (!found) {
            var parts = finalCheck[fc].match(/Page (\d+), (.+) text frame/);
            if (parts) {
                oversetErrors.push({
                    page: parseInt(parts[1], 10),
                    frame: parts[2],
                    location: finalCheck[fc]
                });
            }
        }
    }
    
    if (oversetErrors.length > 0) {
        reportFile.writeln("\n!!! ATTENTION: FOUND " + oversetErrors.length + " OVERSET TEXT ERRORS !!!");
        reportFile.writeln("The following text frames have overset text that is not visible:");
        
        for (var e = 0; e < oversetErrors.length; e++) {
            var error = oversetErrors[e];
            reportFile.writeln("  " + (e+1) + ". " + error.location);
        }
    } else {
        reportFile.writeln("\nGOOD NEWS: No overset text errors found in the document.");
    }
    
    reportFile.writeln("\n=========================\n");
    
    // Add image placement report if automation was used
    if (imageResults.enabled) {
        reportFile.writeln("IMAGE PLACEMENT REPORT:");
        reportFile.writeln("======================");
        reportFile.writeln("Image folder: " + imageResults.folder);
        reportFile.writeln("Total chairpersons processed: " + imageResults.totalAttempted);
        reportFile.writeln("Successfully placed: " + imageResults.successful.length + "/" + imageResults.totalAttempted + " images");
        reportFile.writeln("Missing images: " + imageResults.failed.length + "/" + imageResults.totalAttempted);
        reportFile.writeln("");
        
        if (imageResults.successful.length > 0) {
            reportFile.writeln("SUCCESSFUL PLACEMENTS:");
            for (var s = 0; s < imageResults.successful.length; s++) {
                reportFile.writeln("  [SUCCESS] " + imageResults.successful[s]);
            }
            reportFile.writeln("");
        }
        
        if (imageResults.failed.length > 0) {
            reportFile.writeln("MISSING IMAGES:");
            for (var f = 0; f < imageResults.failed.length; f++) {
                var failed = imageResults.failed[f];
                reportFile.writeln("  [MISSING] " + failed.name);
                if (failed.attempts && failed.attempts.length > 0) {
                    reportFile.writeln("    Searched for: " + failed.attempts.slice(0, 10).join(", "));
                    if (failed.attempts.length > 10) {
                        reportFile.writeln("    ... and " + (failed.attempts.length - 10) + " more variations");
                    }
                }
            }
            reportFile.writeln("");
        }
        
        reportFile.writeln("=========================\n");
    }
    
    reportFile.writeln("SESSIONS REPORT:");
    reportFile.writeln("=========================\n");
    
    for (var i = 0; i < sessionsData.length; i++) {
        var session = sessionsData[i];
        reportFile.writeln("SESSION " + (i+1) + ":");
        reportFile.writeln("  Title: " + (session.title || "[None]"));
        reportFile.writeln("  Time: " + (session.time || "[None]"));
        reportFile.writeln("  Session No: " + (session.no || "[None]"));
        reportFile.writeln("  Chairpersons: " + (session.chairs || "[None]"));
        reportFile.writeln("  Topics: " + (session.topics.length || "0"));
        
        for (var j = 0; j < session.topics.length; j++) {
            var topic = session.topics[j];
            reportFile.writeln("    " + (j+1) + ". " + (topic.title || "[No title]"));
            reportFile.writeln("       Time: " + (topic.time || "[None]"));
            reportFile.writeln("       Speaker: " + (topic.speaker || "[None]"));
        }
        
        reportFile.writeln("");
    }
    
    // Ensure all data is written to the file
    reportFile.write(""); // Force flush
    reportFile.close();
    
    // Verify the file was created successfully
    if (reportFile.exists) {
        $.writeln("Report file created successfully at: " + reportFile.fsName);
    } else {
        $.writeln("Warning: Report file may not have been created properly");
    }
    
    return reportFile; // Return the file object for confirmation
}

/* --------------- RUN --------------- */

try {
    main();
} catch (e) {
    alert("An unexpected error occurred.\n\nError: " + e + (e.line ? ("\nLine: " + e.line) : ""));
}

function copyDocumentSettings(sourceDoc, targetDoc) {
    try {
        // Copy page size and orientation
        targetDoc.documentPreferences.pageWidth = sourceDoc.documentPreferences.pageWidth;
        targetDoc.documentPreferences.pageHeight = sourceDoc.documentPreferences.pageHeight;
        targetDoc.documentPreferences.pageOrientation = sourceDoc.documentPreferences.pageOrientation;
        
        // Copy facing pages setting - CRITICAL for binding
        targetDoc.documentPreferences.facingPages = sourceDoc.documentPreferences.facingPages;
        
        // Copy binding settings - CRITICAL for page direction
        targetDoc.documentPreferences.pageBinding = sourceDoc.documentPreferences.pageBinding;
        targetDoc.documentPreferences.pageDirection = sourceDoc.documentPreferences.pageDirection;
        
        // Copy margins
        targetDoc.documentPreferences.documentMarginTop = sourceDoc.documentPreferences.documentMarginTop;
        targetDoc.documentPreferences.documentMarginBottom = sourceDoc.documentPreferences.documentMarginBottom;
        targetDoc.documentPreferences.documentMarginLeft = sourceDoc.documentPreferences.documentMarginLeft;
        targetDoc.documentPreferences.documentMarginRight = sourceDoc.documentPreferences.documentMarginRight;
        
        // Copy bleed settings
        targetDoc.documentPreferences.documentBleedTopOffset = sourceDoc.documentPreferences.documentBleedTopOffset;
        targetDoc.documentPreferences.documentBleedBottomOffset = sourceDoc.documentPreferences.documentBleedBottomOffset;
        targetDoc.documentPreferences.documentBleedInsideOrLeftOffset = sourceDoc.documentPreferences.documentBleedInsideOrLeftOffset;
        targetDoc.documentPreferences.documentBleedOutsideOrRightOffset = sourceDoc.documentPreferences.documentBleedOutsideOrRightOffset;
        
        // Copy slug settings
        targetDoc.documentPreferences.documentSlugTopOffset = sourceDoc.documentPreferences.documentSlugTopOffset;
        targetDoc.documentPreferences.documentSlugBottomOffset = sourceDoc.documentPreferences.documentSlugBottomOffset;
        targetDoc.documentPreferences.documentSlugInsideOrLeftOffset = sourceDoc.documentPreferences.documentSlugInsideOrLeftOffset;
        targetDoc.documentPreferences.documentSlugOutsideOrRightOffset = sourceDoc.documentPreferences.documentSlugOutsideOrRightOffset;
        
        // Copy page numbering settings
        targetDoc.documentPreferences.pageNumberingStartPage = sourceDoc.documentPreferences.pageNumberingStartPage;
        targetDoc.documentPreferences.pageNumberingSectionPrefix = sourceDoc.documentPreferences.pageNumberingSectionPrefix;
        
        // Copy other document preferences
        targetDoc.documentPreferences.allowPageShuffle = sourceDoc.documentPreferences.allowPageShuffle;
        targetDoc.documentPreferences.allowMasterPageItemsToResize = sourceDoc.documentPreferences.allowMasterPageItemsToResize;
        targetDoc.documentPreferences.allowMasterPageItemsToMove = sourceDoc.documentPreferences.allowMasterPageItemsToMove;
        
        // Copy document intent and output settings
        targetDoc.documentPreferences.intent = sourceDoc.documentPreferences.intent;
        targetDoc.documentPreferences.outputIntent = sourceDoc.documentPreferences.outputIntent;
        
        $.writeln("Document settings copied successfully");
        $.writeln("Facing pages: " + targetDoc.documentPreferences.facingPages);
        $.writeln("Page binding: " + targetDoc.documentPreferences.pageBinding);
        $.writeln("Page direction: " + targetDoc.documentPreferences.pageDirection);
        
    } catch (e) {
        $.writeln("Warning: Could not copy all document settings: " + e);
        $.writeln("Error details: " + e.message);
    }
}