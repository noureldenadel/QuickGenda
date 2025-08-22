// SmartAgenda_Creator.jsx (v2.0 - Styling and Layout)
// Date: August 22, 2025
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
// v2.0 (August 22, 2025) - Bug fixes and improvements
//   • Copy Parent page elements to new document
#target indesign

// Define script version for UI display
var SCRIPT_VERSION = "v2.0 (August 22, 2025)";

// ===================== FORMATTING TAB (COMBINED) =====================
// Top-level definition so it is available to getUnifiedSettingsPanel
function setupFormattingCombinedTab(tab, defaultSettings, hasIndependentLayout, hasTableLayout, availableFields) {
    tab.alignChildren = 'fill';
    tab.margins = [15, 15, 15, 15];

    function getParagraphStyleNames() {
        var names = ['— None —'];
        try {
            var ps = app.activeDocument.paragraphStyles;
            for (var i = 0; i < ps.length; i++) {
                var nm = ps[i].name;
                if (nm && nm !== '[Basic Paragraph]') names.push(nm);
            }
        } catch (e) {}
        return names;
    }
    function getTableStyleNames() {
        var names = ['— None —'];
        try { var ts = app.activeDocument.tableStyles; for (var i=0;i<ts.length;i++) names.push(ts[i].name); } catch(e) {}
        return names;
    }
    function getCellStyleNames() {
        var names = ['— None —'];
        try { var cs = app.activeDocument.cellStyles; for (var i=0;i<cs.length;i++) names.push(cs[i].name); } catch(e) {}
        return names;
    }

    function addStyleRow(parent, labelText, ddItems, preselectText) {
        var grp = parent.add('group'); grp.orientation = 'row'; grp.spacing = 8; grp.alignChildren = 'left';
        var st = grp.add('statictext', undefined, labelText); st.preferredSize.width = 80;
        var dd = grp.add('dropdownlist', undefined, ddItems); dd.preferredSize.width = 220;
        if (preselectText) {
            for (var i=0;i<dd.items.length;i++){ if (dd.items[i].text === preselectText){ dd.selection = i; break; } }
        } else { dd.selection = 0; }
        return { group: grp, label: st, dropdown: dd };
    }
    function addBreakControlsInline(parent, defaultEnabled, defaultChar) {
        var grp = parent.add('group'); grp.orientation = 'row'; grp.spacing = 6; grp.alignChildren = 'left';
        var cb = grp.add('checkbox', undefined, '|->¶'); cb.value = !!defaultEnabled; // visual indicator only
        var st1 = grp.add('statictext', undefined, 'Replace:');
        var et = grp.add('edittext', undefined, (defaultChar || '|')); et.characters = 4;
        var st2 = grp.add('statictext', undefined, 'with line breaks');
        return { group: grp, checkbox: cb, charInput: et };
    }

    // Session Fields
    var pnlSession = tab.add('panel', undefined, 'Session Fields');
    pnlSession.margins = [15, 15, 15, 15]; pnlSession.alignChildren = 'left';
    var paraItems = getParagraphStyleNames();

    var rowTitle = addStyleRow(pnlSession, 'Title:', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.session ? defaultSettings.stylesOptions.session.titlePara : null);
    var brTitle = addBreakControlsInline(rowTitle.group, defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionTitle'] ? defaultSettings.lineBreaks['SessionTitle'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionTitle'] ? defaultSettings.lineBreaks['SessionTitle'].character : '|');

    var rowTime = addStyleRow(pnlSession, 'Time:', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.session ? defaultSettings.stylesOptions.session.timePara : null);
    var brTime = addBreakControlsInline(rowTime.group, defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionTime'] ? defaultSettings.lineBreaks['SessionTime'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionTime'] ? defaultSettings.lineBreaks['SessionTime'].character : '|');

    var rowNo = addStyleRow(pnlSession, 'Number:', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.session ? defaultSettings.stylesOptions.session.noPara : null);
    var brNo = addBreakControlsInline(rowNo.group, defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionNo'] ? defaultSettings.lineBreaks['SessionNo'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionNo'] ? defaultSettings.lineBreaks['SessionNo'].character : '|');

    // Chairpersons
    var pnlChair = tab.add('panel', undefined, 'Chairpersons');
    pnlChair.margins = [15, 15, 15, 15]; pnlChair.alignChildren = 'left';
    var rowChairInline = addStyleRow(pnlChair, 'Style (Inline):', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.chair ? defaultSettings.stylesOptions.chair.inlinePara : null);
    var rowChairGrid = addStyleRow(pnlChair, 'Style (Grid):', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.chair ? defaultSettings.stylesOptions.chair.gridPara : null);
    var brChair = addBreakControlsInline(rowChairInline.group, defaultSettings.lineBreaks && defaultSettings.lineBreaks['Chairpersons'] ? defaultSettings.lineBreaks['Chairpersons'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['Chairpersons'] ? defaultSettings.lineBreaks['Chairpersons'].character : '||');

    // Topics
    var pnlTopics = tab.add('panel', undefined, 'Topics');
    pnlTopics.margins = [15, 15, 15, 15]; pnlTopics.alignChildren = 'left';
    var tableItems = getTableStyleNames();
    var cellItems = getCellStyleNames();
    var rowTbl = addStyleRow(pnlTopics, 'Table:', tableItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.table ? defaultSettings.stylesOptions.table.tableStyle : null);
    var rowCell = addStyleRow(pnlTopics, 'Cells:', cellItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.table ? defaultSettings.stylesOptions.table.cellStyle : null);

    var rowTopicTime = addStyleRow(pnlTopics, 'Time:', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.topicsIndependent ? defaultSettings.stylesOptions.topicsIndependent.timePara : null);
    var brTopicTime = addBreakControlsInline(rowTopicTime.group, defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicTime'] ? defaultSettings.lineBreaks['topicTime'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicTime'] ? defaultSettings.lineBreaks['topicTime'].character : '|');

    var rowTopicTitle = addStyleRow(pnlTopics, 'Title:', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.topicsIndependent ? defaultSettings.stylesOptions.topicsIndependent.titlePara : null);
    var brTopicTitle = addBreakControlsInline(rowTopicTitle.group, defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicTitle'] ? defaultSettings.lineBreaks['topicTitle'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicTitle'] ? defaultSettings.lineBreaks['topicTitle'].character : '|');

    var rowTopicSpeaker = addStyleRow(pnlTopics, 'Speaker:', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.topicsIndependent ? defaultSettings.stylesOptions.topicsIndependent.speakerPara : null);
    var brTopicSpeaker = addBreakControlsInline(rowTopicSpeaker.group, defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicSpeaker'] ? defaultSettings.lineBreaks['topicSpeaker'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicSpeaker'] ? defaultSettings.lineBreaks['topicSpeaker'].character : '|');

    // Create New Style button
    var grpNew = tab.add('group'); grpNew.alignment = 'left';
    var btnNewStyle = grpNew.add('button', undefined, '+ Create New Style...');
    btnNewStyle.onClick = function() {
        try {
            var w = new Window('dialog', 'Create New Style'); w.alignChildren = 'left';
            var g1 = w.add('group'); g1.add('statictext', undefined, 'Name:'); var etName = g1.add('edittext', undefined, ''); etName.characters = 24;
            var g2 = w.add('group'); g2.add('statictext', undefined, 'Type:'); var ddType = g2.add('dropdownlist', undefined, ['Paragraph','Table','Cell']); ddType.selection = 0;
            var gBtn = w.add('group'); var ok = gBtn.add('button', undefined, 'OK'); gBtn.add('button', undefined, 'Cancel', {name:'cancel'});
            ok.onClick = function(){ w.close(1); };
            if (w.show() !== 1) return;
            var name = etName.text; var type = ddType.selection.text;
            if (!name || name === '') return;
            if (type === 'Paragraph') { app.activeDocument.paragraphStyles.add({name:name}); }
            else if (type === 'Table') { app.activeDocument.tableStyles.add({name:name}); }
            else { app.activeDocument.cellStyles.add({name:name}); }
        } catch (e) { alert('Failed to create style: ' + e); }
        // repopulate dropdowns
        var pItems = getParagraphStyleNames(); var tItems = getTableStyleNames(); var cItems = getCellStyleNames();
        function repop(dd, items){ if (!dd) return; var sel = (dd.selection && dd.selection.text) ? dd.selection.text : null; dd.removeAll(); for (var i=0;i<items.length;i++) dd.add('item', items[i]); dd.selection = sel ? (function(){ for (var j=0;j<dd.items.length;j++){ if(dd.items[j].text===sel) return j;} return 0;})() : 0; }
        repop(rowTitle.dropdown, pItems); repop(rowTime.dropdown, pItems); repop(rowNo.dropdown, pItems);
        repop(rowChairInline.dropdown, pItems); repop(rowChairGrid.dropdown, pItems);
        repop(rowTopicTime.dropdown, pItems); repop(rowTopicTitle.dropdown, pItems); repop(rowTopicSpeaker.dropdown, pItems);
        repop(rowTbl.dropdown, tItems); repop(rowCell.dropdown, cItems);
    };

    // Hide rows based on CSV availability
    try {
        if (availableFields) {
            if (!availableFields['Session Title']) { rowTitle.group.visible = false; brTitle.group.visible = false; }
            if (!availableFields['Session Time']) { rowTime.group.visible = false; brTime.group.visible = false; }
            if (!availableFields['Session No']) { rowNo.group.visible = false; brNo.group.visible = false; }
            if (!availableFields['Chairpersons']) { pnlChair.visible = false; }
            if (!availableFields['Time']) { rowTopicTime.group.visible = false; brTopicTime.group.visible = false; }
            if (!availableFields['Topic Title']) { rowTopicTitle.group.visible = false; brTopicTitle.group.visible = false; }
            if (!availableFields['Speaker']) { rowTopicSpeaker.group.visible = false; brTopicSpeaker.group.visible = false; }
        }
    } catch (e) {}

    // Expose properties so existing collectors work
    tab.ddSessionTitle = rowTitle.dropdown;
    tab.ddSessionTime = rowTime.dropdown;
    tab.ddSessionNo = rowNo.dropdown;
    tab.ddChairInline = rowChairInline.dropdown;
    tab.ddChairGrid = rowChairGrid.dropdown;
    tab.ddTableStyle = rowTbl.dropdown;
    tab.ddCellStyle = rowCell.dropdown;
    tab.ddTopicTime = rowTopicTime.dropdown;
    tab.ddTopicTitle = rowTopicTitle.dropdown;
    tab.ddTopicSpeaker = rowTopicSpeaker.dropdown;

    tab.sessionControls = {
        'SessionTitle': { checkbox: brTitle.checkbox, charInput: brTitle.charInput },
        'SessionTime': { checkbox: brTime.checkbox, charInput: brTime.charInput },
        'SessionNo': { checkbox: brNo.checkbox, charInput: brNo.charInput }
    };
    tab.chairpersonControls = {
        'Chairpersons': { checkbox: brChair.checkbox, charInput: brChair.charInput }
    };
    tab.topicControls = {
        'topicTime': { checkbox: brTopicTime.checkbox, charInput: brTopicTime.charInput },
        'topicTitle': { checkbox: brTopicTitle.checkbox, charInput: brTopicTitle.charInput },
        'topicSpeaker': { checkbox: brTopicSpeaker.checkbox, charInput: brTopicSpeaker.charInput }
    };
}

function main() {
    if (app.documents.length === 0) {
        alert("Error: No document is open.\nPlease open your template file first.");
        return;
    }

// ===================== FORMATTING TAB (COMBINED) =====================
function setupFormattingCombinedTab(tab, defaultSettings, hasIndependentLayout, hasTableLayout, availableFields) {
    tab.alignChildren = 'fill';
    tab.margins = [15, 15, 15, 15];

    function getParagraphStyleNames() {
        var names = ['— None —'];
        try {
            var ps = app.activeDocument.paragraphStyles;
            for (var i = 0; i < ps.length; i++) {
                var nm = ps[i].name;
                if (nm && nm !== '[Basic Paragraph]') names.push(nm);
            }
        } catch (e) {}
        return names;
    }
    function getTableStyleNames() {
        var names = ['— None —'];
        try { var ts = app.activeDocument.tableStyles; for (var i=0;i<ts.length;i++) names.push(ts[i].name); } catch(e) {}
        return names;
    }
    function getCellStyleNames() {
        var names = ['— None —'];
        try { var cs = app.activeDocument.cellStyles; for (var i=0;i<cs.length;i++) names.push(cs[i].name); } catch(e) {}
        return names;
    }

    function addStyleRow(parent, labelText, ddItems, preselectText) {
        var grp = parent.add('group'); grp.orientation = 'row'; grp.spacing = 8; grp.alignChildren = 'left';
        var st = grp.add('statictext', undefined, labelText); st.preferredSize.width = 80;
        var dd = grp.add('dropdownlist', undefined, ddItems); dd.preferredSize.width = 220;
        if (preselectText) {
            for (var i=0;i<dd.items.length;i++){ if (dd.items[i].text === preselectText){ dd.selection = i; break; } }
        } else { dd.selection = 0; }
        return { group: grp, label: st, dropdown: dd };
    }
    function addBreakControlsInline(parent, defaultEnabled, defaultChar) {
        var grp = parent.add('group'); grp.orientation = 'row'; grp.spacing = 6; grp.alignChildren = 'left';
        var cb = grp.add('checkbox', undefined, '|->¶'); cb.value = !!defaultEnabled; // visual indicator only
        var st1 = grp.add('statictext', undefined, 'Replace:');
        var et = grp.add('edittext', undefined, (defaultChar || '|')); et.characters = 4;
        var st2 = grp.add('statictext', undefined, 'with line breaks');
        return { group: grp, checkbox: cb, charInput: et };
    }

    // Session Fields
    var pnlSession = tab.add('panel', undefined, 'Session Fields');
    pnlSession.margins = [15, 15, 15, 15]; pnlSession.alignChildren = 'left';
    var paraItems = getParagraphStyleNames();

    var rowTitle = addStyleRow(pnlSession, 'Title:', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.session ? defaultSettings.stylesOptions.session.titlePara : null);
    var brTitle = addBreakControlsInline(rowTitle.group, defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionTitle'] ? defaultSettings.lineBreaks['SessionTitle'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionTitle'] ? defaultSettings.lineBreaks['SessionTitle'].character : '|');

    var rowTime = addStyleRow(pnlSession, 'Time:', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.session ? defaultSettings.stylesOptions.session.timePara : null);
    var brTime = addBreakControlsInline(rowTime.group, defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionTime'] ? defaultSettings.lineBreaks['SessionTime'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionTime'] ? defaultSettings.lineBreaks['SessionTime'].character : '|');

    var rowNo = addStyleRow(pnlSession, 'Number:', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.session ? defaultSettings.stylesOptions.session.noPara : null);
    var brNo = addBreakControlsInline(rowNo.group, defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionNo'] ? defaultSettings.lineBreaks['SessionNo'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionNo'] ? defaultSettings.lineBreaks['SessionNo'].character : '|');

    // Chairpersons
    var pnlChair = tab.add('panel', undefined, 'Chairpersons');
    pnlChair.margins = [15, 15, 15, 15]; pnlChair.alignChildren = 'left';
    var rowChairInline = addStyleRow(pnlChair, 'Style (Inline):', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.chair ? defaultSettings.stylesOptions.chair.inlinePara : null);
    var rowChairGrid = addStyleRow(pnlChair, 'Style (Grid):', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.chair ? defaultSettings.stylesOptions.chair.gridPara : null);
    var brChair = addBreakControlsInline(rowChairInline.group, defaultSettings.lineBreaks && defaultSettings.lineBreaks['Chairpersons'] ? defaultSettings.lineBreaks['Chairpersons'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['Chairpersons'] ? defaultSettings.lineBreaks['Chairpersons'].character : '||');

    // Topics
    var pnlTopics = tab.add('panel', undefined, 'Topics');
    pnlTopics.margins = [15, 15, 15, 15]; pnlTopics.alignChildren = 'left';
    var tableItems = getTableStyleNames();
    var cellItems = getCellStyleNames();
    var rowTbl = addStyleRow(pnlTopics, 'Table:', tableItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.table ? defaultSettings.stylesOptions.table.tableStyle : null);
    var rowCell = addStyleRow(pnlTopics, 'Cells:', cellItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.table ? defaultSettings.stylesOptions.table.cellStyle : null);

    var rowTopicTime = addStyleRow(pnlTopics, 'Time:', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.topicsIndependent ? defaultSettings.stylesOptions.topicsIndependent.timePara : null);
    var brTopicTime = addBreakControlsInline(rowTopicTime.group, defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicTime'] ? defaultSettings.lineBreaks['topicTime'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicTime'] ? defaultSettings.lineBreaks['topicTime'].character : '|');

    var rowTopicTitle = addStyleRow(pnlTopics, 'Title:', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.topicsIndependent ? defaultSettings.stylesOptions.topicsIndependent.titlePara : null);
    var brTopicTitle = addBreakControlsInline(rowTopicTitle.group, defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicTitle'] ? defaultSettings.lineBreaks['topicTitle'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicTitle'] ? defaultSettings.lineBreaks['topicTitle'].character : '|');

    var rowTopicSpeaker = addStyleRow(pnlTopics, 'Speaker:', paraItems, defaultSettings.stylesOptions && defaultSettings.stylesOptions.topicsIndependent ? defaultSettings.stylesOptions.topicsIndependent.speakerPara : null);
    var brTopicSpeaker = addBreakControlsInline(rowTopicSpeaker.group, defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicSpeaker'] ? defaultSettings.lineBreaks['topicSpeaker'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicSpeaker'] ? defaultSettings.lineBreaks['topicSpeaker'].character : '|');

    // Create New Style button
    var grpNew = tab.add('group'); grpNew.alignment = 'left';
    var btnNewStyle = grpNew.add('button', undefined, '+ Create New Style...');
    btnNewStyle.onClick = function() {
        try {
            var w = new Window('dialog', 'Create New Style'); w.alignChildren = 'left';
            var g1 = w.add('group'); g1.add('statictext', undefined, 'Name:'); var etName = g1.add('edittext', undefined, ''); etName.characters = 24;
            var g2 = w.add('group'); g2.add('statictext', undefined, 'Type:'); var ddType = g2.add('dropdownlist', undefined, ['Paragraph','Table','Cell']); ddType.selection = 0;
            var gBtn = w.add('group'); var ok = gBtn.add('button', undefined, 'OK'); gBtn.add('button', undefined, 'Cancel', {name:'cancel'});
            ok.onClick = function(){ w.close(1); };
            if (w.show() !== 1) return;
            var name = etName.text; var type = ddType.selection.text;
            if (!name || name === '') return;
            if (type === 'Paragraph') { app.activeDocument.paragraphStyles.add({name:name}); }
            else if (type === 'Table') { app.activeDocument.tableStyles.add({name:name}); }
            else { app.activeDocument.cellStyles.add({name:name}); }
        } catch (e) { alert('Failed to create style: ' + e); }
        // repopulate dropdowns
        var pItems = getParagraphStyleNames(); var tItems = getTableStyleNames(); var cItems = getCellStyleNames();
        function repop(dd, items){ if (!dd) return; var sel = (dd.selection && dd.selection.text) ? dd.selection.text : null; dd.removeAll(); for (var i=0;i<items.length;i++) dd.add('item', items[i]); dd.selection = sel ? (function(){ for (var j=0;j<dd.items.length;j++){ if(dd.items[j].text===sel) return j;} return 0;})() : 0; }
        repop(rowTitle.dropdown, pItems); repop(rowTime.dropdown, pItems); repop(rowNo.dropdown, pItems);
        repop(rowChairInline.dropdown, pItems); repop(rowChairGrid.dropdown, pItems);
        repop(rowTopicTime.dropdown, pItems); repop(rowTopicTitle.dropdown, pItems); repop(rowTopicSpeaker.dropdown, pItems);
        repop(rowTbl.dropdown, tItems); repop(rowCell.dropdown, cItems);
    };

    // Hide rows based on CSV availability
    try {
        if (availableFields) {
            if (!availableFields['Session Title']) { rowTitle.group.visible = false; brTitle.group.visible = false; }
            if (!availableFields['Session Time']) { rowTime.group.visible = false; brTime.group.visible = false; }
            if (!availableFields['Session No']) { rowNo.group.visible = false; brNo.group.visible = false; }
            if (!availableFields['Chairpersons']) { pnlChair.visible = false; }
            if (!availableFields['Time']) { rowTopicTime.group.visible = false; brTopicTime.group.visible = false; }
            if (!availableFields['Topic Title']) { rowTopicTitle.group.visible = false; brTopicTitle.group.visible = false; }
            if (!availableFields['Speaker']) { rowTopicSpeaker.group.visible = false; brTopicSpeaker.group.visible = false; }
        }
    } catch (e) {}

    // Expose properties so existing collectors work
    tab.ddSessionTitle = rowTitle.dropdown;
    tab.ddSessionTime = rowTime.dropdown;
    tab.ddSessionNo = rowNo.dropdown;
    tab.ddChairInline = rowChairInline.dropdown;
    tab.ddChairGrid = rowChairGrid.dropdown;
    tab.ddTableStyle = rowTbl.dropdown;
    tab.ddCellStyle = rowCell.dropdown;
    tab.ddTopicTime = rowTopicTime.dropdown;
    tab.ddTopicTitle = rowTopicTitle.dropdown;
    tab.ddTopicSpeaker = rowTopicSpeaker.dropdown;

    tab.sessionControls = {
        'SessionTitle': { checkbox: brTitle.checkbox, charInput: brTitle.charInput },
        'SessionTime': { checkbox: brTime.checkbox, charInput: brTime.charInput },
        'SessionNo': { checkbox: brNo.checkbox, charInput: brNo.charInput }
    };
    tab.chairpersonControls = {
        'Chairpersons': { checkbox: brChair.checkbox, charInput: brChair.charInput }
    };
    tab.topicControls = {
        'topicTime': { checkbox: brTopicTime.checkbox, charInput: brTopicTime.charInput },
        'topicTitle': { checkbox: brTopicTitle.checkbox, charInput: brTopicTitle.charInput },
        'topicSpeaker': { checkbox: brTopicSpeaker.checkbox, charInput: brTopicSpeaker.charInput }
    };
}
    var templateDoc = app.activeDocument;

    // --- Create a new document from the template ---
    var doc = app.documents.add();

    // Copy document settings from template to new document BEFORE any page operations
    copyDocumentSettings(templateDoc, doc);

    // Copy master (Parent) spreads from template so applied masters can be preserved
    var masterMap = copyMasterSpreads(templateDoc, doc);

    // Duplicate template pages to the new document (this will be our master template)
    templateDoc.pages.everyItem().duplicate(LocationOptions.AFTER, doc.pages.lastItem());

    // Now remove the original default page (after we have template pages)
    if (doc.pages.length > 1) {
        doc.pages[0].remove();
    }
    // Apply corresponding masters from template to the duplicated pages
    applyMastersToPages(templateDoc, doc, masterMap);

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

    // Build CSV info for CAD Info tab instead of alerting
    var fieldsList = [];
    for (var field in availableFields) {
        if (availableFields[field]) fieldsList.push(field);
    }
    var csvInfo = {
        fieldsList: fieldsList,
        sessionsCount: sessionsData.length,
        version: SCRIPT_VERSION
    };

    // Get all user options from unified settings panel
    var allOptions = getUnifiedSettingsPanel(hasIndependentTopicLayout, hasTableTopicLayout, availableFields, csvInfo);
    if (!allOptions) { doc.close(SaveOptions.NO); return; }

    var templatePage = doc.pages[0]; // This is now our master template page
    for (var i = 0; i < sessionsData.length; i++) {
        // Always duplicate the template page for each session
        var currentPage = templatePage.duplicate(LocationOptions.AFTER, doc.pages.lastItem());
        fillPagePlaceholders(currentPage, sessionsData[i], allOptions.chairOptions, allOptions.topicOptions, allOptions.lineBreakOptions, allOptions.stylesOptions, availableFields);
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

// Create a paragraph style if missing; returns the style name
function createParagraphStyleIfMissing(name, props) {
    try {
        var doc = app.activeDocument;
        var ps = doc.paragraphStyles.itemByName(name);
        if (ps && ps.isValid) return name;
        ps = doc.paragraphStyles.add({ name: name });
        // Apply safe subset of properties
        try { if (props && props.pointSize !== undefined) ps.pointSize = props.pointSize; } catch (e1) {}
        try { if (props && props.leading !== undefined) ps.leading = props.leading; } catch (e2) {}
        try { if (props && props.justification !== undefined) ps.justification = props.justification; } catch (e3) {}
        try { if (props && props.appliedFont) ps.appliedFont = props.appliedFont; } catch (e4) {}
        try { if (props && props.fontStyle) ps.fontStyle = props.fontStyle; } catch (e5) {}
        return name;
    } catch (e) { return name; }
}

// Create a table style if missing; returns the style name
function createTableStyleIfMissing(name, props) {
    try {
        var doc = app.activeDocument;
        var ts = doc.tableStyles.itemByName(name);
        if (ts && ts.isValid) return name;
        ts = doc.tableStyles.add({ name: name });
        // Safe basic props
        try { if (props && props.headerRegionCellStyle) ts.headerRegionCellStyle = props.headerRegionCellStyle; } catch (e1) {}
        try { if (props && props.bodyRegionCellStyle) ts.bodyRegionCellStyle = props.bodyRegionCellStyle; } catch (e2) {}
        return name;
    } catch (e) { return name; }
}

// Create a cell style if missing; returns the style name
function createCellStyleIfMissing(name, props) {
    try {
        var doc = app.activeDocument;
        var cs = doc.cellStyles.itemByName(name);
        if (cs && cs.isValid) return name;
        cs = doc.cellStyles.add({ name: name });
        // Safe basic props
        try { if (props && props.verticalJustification !== undefined) cs.verticalJustification = props.verticalJustification; } catch (e1) {}
        try { if (props && props.topInset !== undefined) cs.topInset = props.topInset; } catch (e2) {}
        try { if (props && props.bottomInset !== undefined) cs.bottomInset = props.bottomInset; } catch (e3) {}
        try { if (props && props.leftInset !== undefined) cs.leftInset = props.leftInset; } catch (e4) {}
        try { if (props && props.rightInset !== undefined) cs.rightInset = props.rightInset; } catch (e5) {}
        return name;
    } catch (e) { return name; }
}

// Ensure a baseline set of default styles exists; returns names used
function ensureDefaultStyles() {
    var defaults = {
        sessionTitle: "SessionTitle",
        sessionTime: "SessionTime",
        tableStyle: "TableStyle",
        headerCell: "HeaderCell",
        bodyCell: "BodyCell"
    };
    try {
        // Paragraph styles
        createParagraphStyleIfMissing(defaults.sessionTitle, { pointSize: 16, leading: 18 });
        createParagraphStyleIfMissing(defaults.sessionTime, { pointSize: 10, leading: 12 });
        // Cell styles first (so we can set them on table style)
        createCellStyleIfMissing(defaults.headerCell, { topInset: 2, bottomInset: 2, leftInset: 2, rightInset: 2 });
        createCellStyleIfMissing(defaults.bodyCell, { topInset: 1, bottomInset: 1, leftInset: 1, rightInset: 1 });
        // Table style referencing cell styles if possible
        var doc = app.activeDocument;
        var headerCS = doc.cellStyles.itemByName(defaults.headerCell);
        var bodyCS = doc.cellStyles.itemByName(defaults.bodyCell);
        createTableStyleIfMissing(defaults.tableStyle, {
            headerRegionCellStyle: headerCS && headerCS.isValid ? headerCS : undefined,
            bodyRegionCellStyle: bodyCS && bodyCS.isValid ? bodyCS : undefined
        });
    } catch (e) {}
    return defaults;
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

// Enumerate Table Styles available in the active document
function getTableStyleNames() {
    var names = [];
    try {
        var doc = app.activeDocument;
        // In some versions, allTableStyles includes defaults; fallback to tableStyles
        var arr = doc.allTableStyles ? doc.allTableStyles : doc.tableStyles;
        for (var i = 0; i < arr.length; i++) {
            try {
                var nm = arr[i].name;
                if (!nm) continue;
                if (nm.charAt(0) === '[') continue; // Skip [Basic Table] etc.
                if (indexOfExact(names, nm) === -1) names.push(nm);
            } catch (ig) {}
        }
    } catch (e) {}
    names.sort();
    return names;
}

// Enumerate Cell Styles available in the active document
function getCellStyleNames() {
    var names = [];
    try {
        var doc = app.activeDocument;
        var arr = doc.allCellStyles ? doc.allCellStyles : doc.cellStyles;
        for (var i = 0; i < arr.length; i++) {
            try {
                var nm = arr[i].name;
                if (!nm) continue;
                if (nm.charAt(0) === '[') continue; // Skip [None] etc.
                if (indexOfExact(names, nm) === -1) names.push(nm);
            } catch (ig) {}
        }
    } catch (e) {}
    names.sort();
    return names;
}
function trimES(s) { if (s === null || s === undefined) return ""; return s.replace(/^\s+|\s+$/g, ""); }

/* ---------------- UNIFIED SETTINGS PANEL ---------------- */

function getUnifiedSettingsPanel(hasIndependentTopicLayout, hasTableTopicLayout, availableFields, csvInfo) {
    var defaultSettings = getDefaultSettings();
    // Ensure a minimal set of default styles exists and prefill defaults if not set
    var createdDefaults = ensureDefaultStyles();
    if (!defaultSettings.stylesOptions) defaultSettings.stylesOptions = {};
    if (!defaultSettings.stylesOptions.session) defaultSettings.stylesOptions.session = {};
    if (!defaultSettings.stylesOptions.table) defaultSettings.stylesOptions.table = {};
    if (!defaultSettings.stylesOptions.chair) defaultSettings.stylesOptions.chair = defaultSettings.stylesOptions.chair || {};
    if (!defaultSettings.stylesOptions.topicsIndependent) defaultSettings.stylesOptions.topicsIndependent = defaultSettings.stylesOptions.topicsIndependent || {};
    // Only set if empty so we don't override user's persisted choices
    if (!defaultSettings.stylesOptions.session.titlePara) defaultSettings.stylesOptions.session.titlePara = createdDefaults.sessionTitle;
    if (!defaultSettings.stylesOptions.session.timePara) defaultSettings.stylesOptions.session.timePara = createdDefaults.sessionTime;
    if (!defaultSettings.stylesOptions.table.tableStyle) defaultSettings.stylesOptions.table.tableStyle = createdDefaults.tableStyle;
    if (!defaultSettings.stylesOptions.table.cellStyle) defaultSettings.stylesOptions.table.cellStyle = createdDefaults.bodyCell;
    var dlg = new Window('dialog', 'SmartAgenda - Configuration');
    dlg.alignChildren = 'fill';
    dlg.preferredSize.width = 600;
    dlg.preferredSize.height = 700;
    
    // Navigation tabs
    var tabGroup = dlg.add('tabbedpanel');
    tabGroup.preferredSize.width = 580;
    tabGroup.preferredSize.height = 600;
    
    // New tab order and names
    var overviewTab = tabGroup.add('tab', undefined, 'Overview');
    var contentTab = tabGroup.add('tab', undefined, 'Content');
    var formattingTab = tabGroup.add('tab', undefined, 'Formatting');
    var advancedTab = tabGroup.add('tab', undefined, 'Advanced');
    
    // Compute template analysis for Overview Dashboard
    var analysisPage = null;
    try { analysisPage = app.activeDocument.pages[0]; } catch (e) { analysisPage = null; }
    var supportsTable = !!hasTableTopicLayout;
    var supportsIndependent = !!hasIndependentTopicLayout;
    var hasAvatar = analysisPage ? !!findItemByLabel(analysisPage, 'chairAvatar') : false;
    var hasFlag = analysisPage ? !!findItemByLabel(analysisPage, 'chairFlag') : false;
    var imageAutomationAvailable = !!(hasAvatar || hasFlag);
    var compatibility = (supportsTable || supportsIndependent) ? 'Good' : 'Limited';
    var suggestedTopicMode = supportsTable ? 'table' : (supportsIndependent ? 'independent' : 'table');
    var suggestedChairMode = 'inline';

    function applySuggested() {
        // Apply topics suggestion
        updateTopicsTabFromSettings(contentTab, { mode: suggestedTopicMode, includeHeader: true });
        // Apply chair suggestion
        updateChairpersonsTabFromSettings(contentTab, { mode: suggestedChairMode });
        // Ensure UI sync
        if (contentTab && typeof contentTab.syncUI === 'function') contentTab.syncUI();
    }

    // Setup Overview (formerly Info) as Dashboard
    setupCADInfoTab(overviewTab, csvInfo, {
        supportsTable: supportsTable,
        supportsIndependent: supportsIndependent,
        imageAutomationAvailable: imageAutomationAvailable,
        compatibility: compatibility,
        suggested: {
            topicMode: suggestedTopicMode,
            chairMode: suggestedChairMode
        }
    }, applySuggested);
    
    // Setup Content (merge Chairpersons + Topics)
    setupChairpersonsTab(contentTab, defaultSettings);
    setupTopicsTab(contentTab, defaultSettings, hasIndependentTopicLayout, hasTableTopicLayout);
    
    // Setup Formatting (merged Line Breaks + Styles)
    setupFormattingCombinedTab(
        formattingTab,
        defaultSettings,
        hasIndependentTopicLayout,
        hasTableTopicLayout,
        availableFields
    );
    
    // ===== ADVANCED: CENTRALIZED SETTINGS MANAGEMENT =====
    // Top strip: Import/Export all (entire configuration)
    var pnlSettings = advancedTab.add('panel', undefined, 'Settings Management');
    pnlSettings.orientation = 'row';
    pnlSettings.alignChildren = 'center';
    pnlSettings.margins = [15, 15, 15, 15];
    var btnImport = pnlSettings.add('button', undefined, 'Import All Settings');
    var btnExport = pnlSettings.add('button', undefined, 'Export All Settings');
    btnImport.preferredSize.width = 150;
    btnExport.preferredSize.width = 150;

    // Profiles panel: save/load subsets of settings
    var pnlProfiles = advancedTab.add('panel', undefined, 'Profiles');
    pnlProfiles.orientation = 'column';
    pnlProfiles.alignChildren = 'left';
    pnlProfiles.margins = [15, 15, 15, 15];

    var grpProfileTop = pnlProfiles.add('group');
    grpProfileTop.orientation = 'row';
    grpProfileTop.spacing = 10;
    grpProfileTop.add('statictext', undefined, 'Profile:').preferredSize.width = 60;
    var ddProfiles = grpProfileTop.add('dropdownlist', undefined, []);
    ddProfiles.preferredSize.width = 220;
    var btnRefreshProfiles = grpProfileTop.add('button', undefined, 'Refresh');
    var btnDeleteProfile = grpProfileTop.add('button', undefined, 'Delete');

    var grpProfileActions = pnlProfiles.add('group');
    grpProfileActions.orientation = 'row';
    grpProfileActions.spacing = 10;
    var btnSaveProfile = grpProfileActions.add('button', undefined, 'Save As...');
    var btnLoadProfile = grpProfileActions.add('button', undefined, 'Load');
    var btnResetDefaults = grpProfileActions.add('button', undefined, 'Reset Defaults');

    // Inclusion options panel for profiles
    var pnlInclude = advancedTab.add('panel', undefined, 'Profile Content');
    pnlInclude.orientation = 'column';
    pnlInclude.alignChildren = 'left';
    pnlInclude.margins = [15, 15, 15, 15];
    var cbIncChair = pnlInclude.add('checkbox', undefined, 'Include Chairpersons settings'); cbIncChair.value = true;
    var cbIncTopics = pnlInclude.add('checkbox', undefined, 'Include Topics settings'); cbIncTopics.value = true;
    var cbIncLineBreaks = pnlInclude.add('checkbox', undefined, 'Include Line Break rules'); cbIncLineBreaks.value = true;
    var cbIncStyles = pnlInclude.add('checkbox', undefined, 'Include Styles selections'); cbIncStyles.value = true;

    // Debugging & Troubleshooting
    var pnlDebug = advancedTab.add('panel', undefined, 'Debugging & Troubleshooting');
    pnlDebug.orientation = 'column';
    pnlDebug.alignChildren = 'left';
    pnlDebug.margins = [15, 15, 15, 15];
    var cbVerbose = pnlDebug.add('checkbox', undefined, 'Verbose console logging ($.writeln)'); cbVerbose.value = false;
    var cbWarnAsAlerts = pnlDebug.add('checkbox', undefined, 'Show warnings as alerts'); cbWarnAsAlerts.value = false;
    var cbValidatePlaceholders = pnlDebug.add('checkbox', undefined, 'Validate template placeholders on start'); cbValidatePlaceholders.value = true;
    var cbStopOnError = pnlDebug.add('checkbox', undefined, 'Stop on first error'); cbStopOnError.value = false;

    // Report preferences
    var pnlReport = advancedTab.add('panel', undefined, 'Report Preferences');
    pnlReport.orientation = 'column';
    pnlReport.alignChildren = 'left';
    pnlReport.margins = [15, 15, 15, 15];
    var cbAutoOpenReport = pnlReport.add('checkbox', undefined, 'Auto-open report after generation'); cbAutoOpenReport.value = true;
    var cbReportImages = pnlReport.add('checkbox', undefined, 'Include image placement results'); cbReportImages.value = true;
    var cbReportOverset = pnlReport.add('checkbox', undefined, 'Include overset text summary'); cbReportOverset.value = true;
    var cbReportCounts = pnlReport.add('checkbox', undefined, 'Include session/topic counts'); cbReportCounts.value = true;

    // Utility: Profiles storage under user data folder
    function getProfilesFolder() {
        try {
            var base = Folder.userData;
            var folder = new Folder(base.fsName + '/SmartAgenda_Profiles');
            if (!folder.exists) folder.create();
            return folder;
        } catch (e) {
            return Folder.desktop; // fallback
        }
    }
    function listProfiles() {
        var folder = getProfilesFolder();
        var files = folder.getFiles(function(f){ return f instanceof File && /\.json$/i.test(f.name); });
        var names = [];
        for (var i=0;i<files.length;i++) names.push(files[i].name.replace(/\.json$/i, ''));
        names.sort();
        return names;
    }
    function refreshProfilesDD(selectName) {
        ddProfiles.removeAll();
        var names = listProfiles();
        for (var i=0;i<names.length;i++) ddProfiles.add('item', names[i]);
        if (names.length > 0) {
            var idx = selectName ? indexOfExact(names, selectName) : 0;
            ddProfiles.selection = (idx >= 0 ? idx : 0);
        }
    }
    function readProfileByName(name) {
        var file = new File(getProfilesFolder().fsName + '/' + name + '.json');
        if (!file.exists) return null;
        try { file.encoding = 'UTF-8'; file.open('r'); var txt = file.read(); file.close(); return jsonFromString(txt); } catch (e) { return null; }
    }
    function saveProfileByName(name, data) {
        var file = new File(getProfilesFolder().fsName + '/' + name + '.json');
        try { file.encoding = 'UTF-8'; file.open('w'); file.write(settingsToJSON(data)); file.close(); return true; } catch (e) { alert('Failed to save profile: ' + e); return false; }
    }
    function deleteProfileByName(name) {
        var file = new File(getProfilesFolder().fsName + '/' + name + '.json');
        if (file.exists) { try { return file.remove(); } catch (e) { return false; } }
        return false;
    }

    // Wire Import/Export All handlers
    btnImport.onClick = function() {
        var importedSettings = importAllSettings();
        if (importedSettings) {
            updateChairpersonsTabFromSettings(contentTab, importedSettings.chairOptions || {});
            updateTopicsTabFromSettings(contentTab, importedSettings.topicOptions || {});
            updateLineBreaksTabFromSettings(formattingTab, importedSettings.lineBreakOptions || {});
            updateStylesTabFromSettings(formattingTab, (importedSettings.stylesOptions || {}), hasIndependentTopicLayout, hasTableTopicLayout, availableFields);
            alert('All settings imported successfully!');
        }
    };
    btnExport.onClick = function() {
        var allSettings = collectAllSettings(contentTab, contentTab, formattingTab, formattingTab);
        if (exportAllSettings(allSettings)) { alert('All settings exported successfully!'); }
    };

    // Profiles: Save
    btnSaveProfile.onClick = function() {
        var all = collectAllSettings(contentTab, contentTab, formattingTab, formattingTab);
        var subset = {};
        if (cbIncChair.value) subset.chairOptions = all.chairOptions;
        if (cbIncTopics.value) subset.topicOptions = all.topicOptions;
        if (cbIncLineBreaks.value) subset.lineBreakOptions = all.lineBreakOptions;
        if (cbIncStyles.value) subset.stylesOptions = all.stylesOptions;

        var w = new Window('dialog', 'Save Profile');
        w.alignChildren = 'left';
        w.margins = [15,15,15,15];
        var grp = w.add('group'); grp.orientation = 'row'; grp.spacing = 10;
        grp.add('statictext', undefined, 'Name:');
        var et = grp.add('edittext', undefined, 'MyProfile'); et.characters = 24;
        var gbtn = w.add('group'); gbtn.orientation = 'row'; gbtn.alignment = 'right';
        var ok = gbtn.add('button', undefined, 'Save'); var cancel = gbtn.add('button', undefined, 'Cancel', {name:'cancel'});
        if (w.show() === 1) {
            var name = trimES(et.text);
            if (!name) { alert('Please enter a profile name.'); return; }
            if (saveProfileByName(name, subset)) {
                refreshProfilesDD(name);
                alert('Profile saved.');
            }
        }
    };

    // Profiles: Load
    btnLoadProfile.onClick = function() {
        if (!ddProfiles.selection) { alert('No profile selected.'); return; }
        var name = ddProfiles.selection.text;
        var data = readProfileByName(name);
        if (!data) { alert('Failed to load profile.'); return; }
        if (data.chairOptions && cbIncChair.value) updateChairpersonsTabFromSettings(contentTab, data.chairOptions);
        if (data.topicOptions && cbIncTopics.value) updateTopicsTabFromSettings(contentTab, data.topicOptions);
        if (data.lineBreakOptions && cbIncLineBreaks.value) updateLineBreaksTabFromSettings(formattingTab, data.lineBreakOptions);
        if (data.stylesOptions && cbIncStyles.value) updateStylesTabFromSettings(formattingTab, data.stylesOptions, hasIndependentTopicLayout, hasTableTopicLayout, availableFields);
        alert('Profile loaded.');
    };

    // Profiles: Delete
    btnDeleteProfile.onClick = function() {
        if (!ddProfiles.selection) { alert('No profile selected.'); return; }
        var name = ddProfiles.selection.text;
        if (confirm('Delete profile "' + name + '"?')) {
            if (deleteProfileByName(name)) { refreshProfilesDD(); }
            else { alert('Failed to delete profile.'); }
        }
    };

    // Profiles: Refresh
    btnRefreshProfiles.onClick = function(){ refreshProfilesDD(ddProfiles.selection ? ddProfiles.selection.text : null); };

    // Reset defaults: re-apply initial defaults via existing updaters
    btnResetDefaults.onClick = function() {
        var defs = getDefaultSettings();
        updateChairpersonsTabFromSettings(contentTab, {
            mode: defs.chairMode,
            inlineSeparator: defs.chairInlineSeparator,
            order: defs.chairOrder,
            columns: defs.chairColumns,
            rows: defs.chairRows,
            colSpacing: formatValueWithUnit(defs.chairColSpacing, defs.chairUnits),
            rowSpacing: formatValueWithUnit(defs.chairRowSpacing, defs.chairUnits),
            units: defs.chairUnits,
            centerGrid: defs.chairCenterGrid
        });
        updateTopicsTabFromSettings(contentTab, {
            mode: defs.topicMode || (hasTableTopicLayout ? 'table' : 'independent'),
            includeHeader: defs.topicIncludeHeader,
            verticalSpacing: '' + defs.topicVerticalSpacing,
            units: defs.topicUnits
        });
        updateLineBreaksTabFromSettings(formattingTab, { lineBreaks: defs.lineBreaks });
        updateStylesTabFromSettings(formattingTab, defs.stylesOptions || {}, hasIndependentTopicLayout, hasTableTopicLayout, availableFields);
        if (typeof contentTab.syncUI === 'function') contentTab.syncUI();
        alert('Defaults restored.');
    };

    // Expose advanced controls for potential future use
    advancedTab.profileDropdown = ddProfiles;
    advancedTab.includeChair = cbIncChair;
    advancedTab.includeTopics = cbIncTopics;
    advancedTab.includeLineBreaks = cbIncLineBreaks;
    advancedTab.includeStyles = cbIncStyles;
    advancedTab.debugVerbose = cbVerbose;
    advancedTab.debugWarnAsAlerts = cbWarnAsAlerts;
    advancedTab.debugValidate = cbValidatePlaceholders;
    advancedTab.debugStopOnError = cbStopOnError;
    advancedTab.reportAutoOpen = cbAutoOpenReport;
    advancedTab.reportIncludeImages = cbReportImages;
    advancedTab.reportIncludeOverset = cbReportOverset;
    advancedTab.reportIncludeCounts = cbReportCounts;

    // Initialize profiles list
    refreshProfilesDD();
    
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
    
    // Collect all settings and return using merged tabs
    return collectAllSettings(contentTab, contentTab, formattingTab, formattingTab);
}

// CAD Info Tab: displays version info and CSV detection summary
function setupCADInfoTab(tab, csvInfo, analysis, onApplySuggested) {
    tab.alignChildren = 'left';
    tab.margins = [15, 15, 15, 15];

    // CSV File Analysis
    var pnlCSV = tab.add('panel', undefined, 'CSV File Analysis');
    pnlCSV.margins = [15, 15, 15, 15];
    pnlCSV.alignChildren = 'left';
    var grpSessions = pnlCSV.add('group');
    grpSessions.orientation = 'row';
    grpSessions.spacing = 10;
    grpSessions.alignment = 'left';
    var stSessLabel = grpSessions.add('statictext', undefined, 'Sessions found:');
    stSessLabel.preferredSize.width = 160;
    var stSessVal = grpSessions.add('statictext', undefined, '' + ((csvInfo && csvInfo.sessionsCount != null) ? csvInfo.sessionsCount : 0));
    stSessVal.alignment = 'left';

    var grpCompat = pnlCSV.add('group');
    grpCompat.orientation = 'row';
    grpCompat.spacing = 10;
    grpCompat.alignment = 'left';
    var stCompatLabel = grpCompat.add('statictext', undefined, 'Template compatibility:');
    stCompatLabel.preferredSize.width = 160;
    var stCompatVal = grpCompat.add('statictext', undefined, (analysis && analysis.compatibility) ? analysis.compatibility : 'Unknown');
    stCompatVal.alignment = 'left';

    // Detected fields list (bigger to avoid scrollbars/overflow)
    var grpFields = pnlCSV.add('group');
    grpFields.orientation = 'column';
    grpFields.alignChildren = 'left';
    grpFields.alignment = 'left';
    grpFields.spacing = 6;
    grpFields.add('statictext', undefined, 'Detected fields:').alignment = 'left';
    var lbFields = grpFields.add('listbox', undefined, (csvInfo && csvInfo.fieldsList) ? csvInfo.fieldsList : []);
    lbFields.preferredSize.width = 520;
    lbFields.preferredSize.height = 220;
    lbFields.alignment = 'left';

    // Template Analysis
    var pnlTemplate = tab.add('panel', undefined, 'Template Analysis');
    pnlTemplate.margins = [15, 15, 15, 15];
    pnlTemplate.alignChildren = 'left';
    function yn(vTrue) { return vTrue ? 'Yes' : 'No'; }
    var g1 = pnlTemplate.add('group'); g1.orientation = 'row'; g1.spacing = 10; g1.alignment = 'left'; var g1l = g1.add('statictext', undefined, 'Supports table layout:'); g1l.preferredSize.width = 200; var g1v = g1.add('statictext', undefined, yn(analysis && analysis.supportsTable)); g1v.alignment = 'left';
    var g2 = pnlTemplate.add('group'); g2.orientation = 'row'; g2.spacing = 10; g2.alignment = 'left'; var g2l = g2.add('statictext', undefined, 'Supports independent layout:'); g2l.preferredSize.width = 200; var g2v = g2.add('statictext', undefined, yn(analysis && analysis.supportsIndependent)); g2v.alignment = 'left';
    var g3 = pnlTemplate.add('group'); g3.orientation = 'row'; g3.spacing = 10; g3.alignment = 'left'; var g3l = g3.add('statictext', undefined, 'Image automation available:'); g3l.preferredSize.width = 200; var g3v = g3.add('statictext', undefined, yn(analysis && analysis.imageAutomationAvailable)); g3v.alignment = 'left';

    // Suggested Configuration (no Apply button)
    var pnlSuggest = tab.add('panel', undefined, 'Suggested Configuration');
    pnlSuggest.margins = [15, 15, 15, 15];
    pnlSuggest.alignChildren = 'left';
    var s1 = pnlSuggest.add('group'); s1.orientation = 'row'; s1.spacing = 10; s1.alignment = 'left'; var s1l = s1.add('statictext', undefined, 'Topic layout:'); s1l.preferredSize.width = 160; var s1v = s1.add('statictext', undefined, analysis && analysis.suggested ? (analysis.suggested.topicMode === 'table' ? 'Table' : 'Independent') : '—'); s1v.alignment = 'left';
    var s2 = pnlSuggest.add('group'); s2.orientation = 'row'; s2.spacing = 10; s2.alignment = 'left'; var s2l = s2.add('statictext', undefined, 'Chairpersons layout:'); s2l.preferredSize.width = 160; var s2v = s2.add('statictext', undefined, analysis && analysis.suggested ? (analysis.suggested.chairMode === 'inline' ? 'Inline' : 'Grid') : '—'); s2v.alignment = 'left';

    // Script info panel
    var pnlVersion = tab.add('panel', undefined, 'Script Information');
    pnlVersion.margins = [15, 15, 15, 15];
    pnlVersion.alignChildren = 'left';
    var txtVer = pnlVersion.add('statictext', undefined, 'Version: ' + (csvInfo && csvInfo.version ? csvInfo.version : 'Unknown'));
    txtVer.graphics.font = ScriptUI.newFont(txtVer.graphics.font.name, ScriptUI.FontStyle.ITALIC, 11);
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
    // Add radios for clearer layout selection and hide the dropdown (kept for data collection)
    var grpRadios = pnlDesc.add('group');
    grpRadios.orientation = 'row';
    grpRadios.spacing = 20;
    var rbInline = grpRadios.add('radiobutton', undefined, 'Inline');
    var rbGrid = grpRadios.add('radiobutton', undefined, 'Grid');
    ddMode.visible = false;
    // Initialize selection based on defaults
    try {
        if (defaultSettings && defaultSettings.chairMode === 'grid') {
            ddMode.selection = 1; rbGrid.value = true;
        } else { ddMode.selection = 0; rbInline.value = true; }
    } catch (e) { ddMode.selection = 0; rbInline.value = true; }
    // Keep ddMode selection in sync for existing collectors
    rbInline.onClick = function(){ ddMode.selection = 0; if (typeof tab.syncUI === 'function') tab.syncUI(); };
    rbGrid.onClick = function(){ ddMode.selection = 1; if (typeof tab.syncUI === 'function') tab.syncUI(); };
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
    
    var grpSpacing = pnlGrid.add('group');
    grpSpacing.orientation = 'column';
    grpSpacing.alignChildren = 'left';
    grpSpacing.margins = [0, 5, 0, 0];

    var grpCol = grpSpacing.add('group'); grpCol.orientation = 'row'; grpCol.spacing = 6; grpCol.alignment = 'left';
    grpCol.add('statictext', undefined, 'Col spacing:').preferredSize.width = 80;
    var etColSpace = grpCol.add('edittext', undefined, formatValueWithUnit(defaultSettings.chairColSpacing || 8, defaultSettings.chairUnits || 'pt')); etColSpace.characters = 8;
    var btnColMinus = grpCol.add('button', undefined, '-'); btnColMinus.preferredSize.width = 22;
    var btnColPlus = grpCol.add('button', undefined, '+'); btnColPlus.preferredSize.width = 22;

    var grpRow = grpSpacing.add('group'); grpRow.orientation = 'row'; grpRow.spacing = 6; grpRow.alignment = 'left';
    grpRow.add('statictext', undefined, 'Row spacing:').preferredSize.width = 80;
    var etRowSpace = grpRow.add('edittext', undefined, formatValueWithUnit(defaultSettings.chairRowSpacing || 4, defaultSettings.chairUnits || 'pt')); etRowSpace.characters = 8;
    var btnRowMinus = grpRow.add('button', undefined, '-'); btnRowMinus.preferredSize.width = 22;
    var btnRowPlus = grpRow.add('button', undefined, '+'); btnRowPlus.preferredSize.width = 22;

    // Stepper handlers
    btnColMinus.onClick = function(){ etColSpace.text = stepValueWithUnit(etColSpace.text, -1, defaultSettings.chairUnits || 'pt'); };
    btnColPlus.onClick = function(){ etColSpace.text = stepValueWithUnit(etColSpace.text, +1, defaultSettings.chairUnits || 'pt'); };
    btnRowMinus.onClick = function(){ etRowSpace.text = stepValueWithUnit(etRowSpace.text, -1, defaultSettings.chairUnits || 'pt'); };
    btnRowPlus.onClick = function(){ etRowSpace.text = stepValueWithUnit(etRowSpace.text, +1, defaultSettings.chairUnits || 'pt'); };
    
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
    tab.cbCenter = cbCenter;
    tab.cbEnableImages = cbEnableImages;
    tab.etImageFolder = etImageFolder;
    tab.ddImageFitting = ddImageFitting;
    
    // Make syncUI accessible from outside
    tab.syncUI = function() {
        var gridMode = ddMode.selection.index === 1;
        // Contextual show/hide to reduce clutter
        if (pnlInline) { pnlInline.enabled = !gridMode; pnlInline.visible = !gridMode; }
        if (pnlGrid) { pnlGrid.enabled = gridMode; pnlGrid.visible = gridMode; }
        if (gridMode) {
            var rowFirst = ddOrder.selection.index === 0;
            if (etCols) etCols.enabled = rowFirst;
            if (etRows) etRows.enabled = !rowFirst;
        }
        // Sync image controls - only available in grid mode
        syncImageUI();
    };

    // Lightweight input validation and sanitization
    try {
        if (etCols && typeof etCols.onChanging === 'undefined') {
            etCols.onChanging = function(){ this.text = this.text.replace(/[^0-9]/g, ''); };
        }
        if (etRows && typeof etRows.onChanging === 'undefined') {
            etRows.onChanging = function(){ this.text = this.text.replace(/[^0-9]/g, ''); };
        }
        if (etColSpace && typeof etColSpace.onChanging === 'undefined') {
            etColSpace.onChanging = function(){ var p = parseValueWithUnit(this.text, defaultSettings.chairUnits || 'pt'); this.text = formatValueWithUnit(p.value, p.unit); };
        }
        if (etRowSpace && typeof etRowSpace.onChanging === 'undefined') {
            etRowSpace.onChanging = function(){ var p = parseValueWithUnit(this.text, defaultSettings.chairUnits || 'pt'); this.text = formatValueWithUnit(p.value, p.unit); };
        }
    } catch (e) {}
    
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
        // Contextual show/hide to reduce clutter
        pnlTable.visible = tableMode;
        pnlIndependent.visible = !tableMode;
    }
    
    rbTable.onClick = syncUI;
    rbIndependent.onClick = syncUI;
    // Inline validation for spacing
    try { if (etSpacing && typeof etSpacing.onChanging === 'undefined') { etSpacing.onChanging = function(){ this.text = this.text.replace(/[^0-9]/g, ''); }; } } catch (e) {}
    // Show reason if a mode is disabled
    try {
        if (!hasTableLayout) {
            var note = pnlMode.add('statictext', undefined, 'Table layout disabled: template has no topicsTable label.');
            note.graphics.foregroundColor = note.graphics.newPen(note.graphics.PenType.SOLID_COLOR, [0.6,0.2,0.2], 1);
        }
        if (!hasIndependentLayout) {
            var note2 = pnlMode.add('statictext', undefined, 'Independent rows disabled: template missing topicTime/title/speaker frames.');
            note2.graphics.foregroundColor = note2.graphics.newPen(note2.graphics.PenType.SOLID_COLOR, [0.6,0.2,0.2], 1);
        }
    } catch (e) {}
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

function collectAllSettings(chairTab, topicTab, lineBreakTab, stylesTab) {
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

    // Collect styles (Phase 1 - paragraph styles)
    function val(dd) { return (dd && dd.selection) ? (dd.selection.index === 0 ? '' : dd.selection.text) : ''; }
    var stylesOptions = {
        table: {
            tableStyle: val(stylesTab && stylesTab.ddTableStyle),
            cellStyle: val(stylesTab && stylesTab.ddCellStyle)
        },
        session: {
            titlePara: val(stylesTab && stylesTab.ddSessionTitle),
            timePara: val(stylesTab && stylesTab.ddSessionTime),
            noPara: val(stylesTab && stylesTab.ddSessionNo)
        },
        chair: {
            inlinePara: val(stylesTab && stylesTab.ddChairInline),
            gridPara: val(stylesTab && stylesTab.ddChairGrid)
        },
        topicsIndependent: {
            timePara: val(stylesTab && stylesTab.ddTopicTime),
            titlePara: val(stylesTab && stylesTab.ddTopicTitle),
            speakerPara: val(stylesTab && stylesTab.ddTopicSpeaker)
        }
    };

    return {
        chairOptions: chairOptions,
        topicOptions: topicOptions,
        lineBreakOptions: lineBreakOptions,
        stylesOptions: stylesOptions
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

// Parse strings like "8pt", "5 mm", "10.5cm" => { value: 8, unit: 'pt' }
function parseValueWithUnit(txt, defaultUnit) {
    if (!txt) return { value: 0, unit: defaultUnit || 'pt' };
    var m = ("" + txt).trim().match(/^\s*([+-]?\d*(?:\.\d+)?)\s*([a-zA-Z%]*)\s*$/);
    var unit = defaultUnit || 'pt';
    var val = 0;
    if (m) {
        val = parseFloat(m[1]);
        if (isNaN(val)) val = 0;
        if (m[2]) unit = m[2].toLowerCase();
    }
    if (unit !== 'pt' && unit !== 'mm' && unit !== 'cm' && unit !== 'px') unit = defaultUnit || 'pt';
    return { value: val, unit: unit };
}

function formatValueWithUnit(value, unit) {
    var v = (typeof value === 'number' && !isNaN(value)) ? value : parseFloat(value);
    if (isNaN(v)) v = 0;
    var u = unit || 'pt';
    return (Math.round(v * 100) / 100) + u;
}

function stepValueWithUnit(current, delta, defaultUnit) {
    var p = parseValueWithUnit(current, defaultUnit);
    var step = 1;
    // Use smaller step for non-pt metric units
    if (p.unit === 'mm') step = 0.5;
    if (p.unit === 'cm') step = 0.1;
    if (p.unit === 'px') step = 1;
    var nv = p.value + (delta > 0 ? step : -step);
    if (nv < 0) nv = 0;
    return formatValueWithUnit(nv, p.unit);
}

// Panel styling for visual hierarchy
// level: 'primary' | 'secondary' | 'option'
function applyPanelStyle(panel, level) {
    try {
        var g = panel.graphics;
        var bg = [0.92, 0.92, 0.92];
        var border = [0.60, 0.60, 0.60];
        if (level === 'primary') { bg = [0.85, 0.85, 0.85]; border = [0.30, 0.30, 0.30]; }
        else if (level === 'secondary') { bg = [0.90, 0.90, 0.90]; border = [0.45, 0.45, 0.45]; }
        else if (level === 'option') { bg = [0.96, 0.96, 0.96]; border = [0.75, 0.75, 0.75]; }
        g.backgroundColor = g.newBrush(g.BrushType.SOLID_COLOR, bg);
        g.pen = g.newPen(g.PenType.SOLID_COLOR, border, 1);
    } catch (e) {}
}

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

function fillPagePlaceholders(page, session, chairOptions, topicOptions, lineBreakOptions, stylesOptions, availableFields) {
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
        populateTopicsAsTable(placeholders.topicsTable, session.topics, topicOptions, lineBreakOptions, stylesOptions);
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

function populateTopicsAsTable(tableFrame, topics, opts, lineBreakOptions, stylesOptions) {
    // Determine desired structure
    var desiredCols = 3; // Time, Topic, Speaker
    var includeHeader = !!opts.includeHeader;
    var desiredHeaderRows = includeHeader ? 1 : 0;

    // Find or create table without destroying existing styling
    var tbl = null;
    if (tableFrame.tables.length > 0) {
        tbl = tableFrame.tables[0];
    } else {
        // Create a fresh table if none exists
        if (tableFrame.contents === undefined) return; // safety
        try { tableFrame.contents = ""; } catch (e) {}
        tbl = tableFrame.texts[0].tables.add({
            bodyRowCount: Math.max(1, topics.length),
            columnCount: desiredCols
        });
    }

    // Set column widths proportionally to the frame width
    try {
        var gb = tableFrame.geometricBounds;
        var frameWidth = gb[3] - gb[1];
        if (tbl.columns.length >= 3) {
            tbl.columns[0].width = frameWidth * 0.20;
            tbl.columns[1].width = frameWidth * 0.50;
            tbl.columns[2].width = frameWidth * 0.30;
        }
    } catch (e) {}

    // Apply Table and Cell Styles if specified
    try {
        var doc = app.activeDocument;
        if (stylesOptions && stylesOptions.table) {
            var tsName = stylesOptions.table.tableStyle || '';
            var csName = stylesOptions.table.cellStyle || '';
            if (tsName) {
                try {
                    var tStyle = doc.tableStyles.itemByName(tsName);
                    if (tStyle && tStyle.isValid) tbl.appliedTableStyle = tStyle;
                } catch (eTS) {}
            }
            if (csName) {
                try {
                    var cStyle = doc.cellStyles.itemByName(csName);
                    if (cStyle && cStyle.isValid) {
                        tbl.cells.everyItem().appliedCellStyle = cStyle;
                    }
                } catch (eCS) {}
            }
        }
    } catch (eApply) {}

    // Ensure header row count matches option (converts first row to header if needed)
    try { tbl.headerRowCount = desiredHeaderRows; } catch (e) {}

    // Ensure column count is exactly desiredCols (add/remove at end)
    try {
        while (tbl.columns.length < desiredCols) tbl.columns.add();
        while (tbl.columns.length > desiredCols) tbl.columns.lastItem().remove();
    } catch (e) {}

    // Calculate required total rows
    var requiredTotalRows = desiredHeaderRows + Math.max(0, topics.length);

    // Add/remove rows at the end to match required count
    try {
        while (tbl.rows.length < requiredTotalRows) tbl.rows.add();
        while (tbl.rows.length > requiredTotalRows && tbl.rows.length > 0) tbl.rows.lastItem().remove();
    } catch (e) {}

    // If header enabled, set header labels (contents only; preserves cell styles)
    if (includeHeader && tbl.rows.length >= 1 && tbl.columns.length >= 3) {
        try {
            tbl.rows[0].cells[0].contents = "Time";
            tbl.rows[0].cells[1].contents = "Topic";
            tbl.rows[0].cells[2].contents = "Speaker";
        } catch (e) {}
    }

    // If no topics, we are done after resizing (header may remain)
    if (topics.length === 0) return;

    // Fill body rows cell-by-cell without altering cell/table styles
    var startRowIndex = desiredHeaderRows; // first body row index
    for (var i = 0; i < topics.length; i++) {
        var rIndex = startRowIndex + i;
        // Safety checks
        if (rIndex >= tbl.rows.length) break;
        if (tbl.columns.length < 3) break;

        // Time
        try {
            tbl.rows[rIndex].cells[0].contents = topics[i].time || "";
            if (lineBreakOptions.lineBreaks.topicTime && lineBreakOptions.lineBreaks.topicTime.enabled) {
                applyLineBreaksToCell(tbl.rows[rIndex].cells[0], lineBreakOptions.lineBreaks.topicTime.character);
            }
        } catch (e) {}

        // Topic
        try {
            tbl.rows[rIndex].cells[1].contents = topics[i].title || "";
            if (lineBreakOptions.lineBreaks.topicTitle && lineBreakOptions.lineBreaks.topicTitle.enabled) {
                applyLineBreaksToCell(tbl.rows[rIndex].cells[1], lineBreakOptions.lineBreaks.topicTitle.character);
            }
        } catch (e) {}

        // Speaker
        try {
            tbl.rows[rIndex].cells[2].contents = topics[i].speaker || "";
            if (lineBreakOptions.lineBreaks.topicSpeaker && lineBreakOptions.lineBreaks.topicSpeaker.enabled) {
                applyLineBreaksToCell(tbl.rows[rIndex].cells[2], lineBreakOptions.lineBreaks.topicSpeaker.character);
            }
        } catch (e) {}
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
        // Styles (Phase 1 - paragraph styles only)
        stylesOptions: {
            table: {
                tableStyle: '',
                cellStyle: ''
            },
            session: {
                titlePara: '',
                timePara: '',
                noPara: ''
            },
            chair: {
                inlinePara: '',
                gridPara: ''
            },
            topicsIndependent: {
                timePara: '',
                titlePara: '',
                speakerPara: ''
            }
        },
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
                        location: "Page " + (p+1) + ", " + itemLabel + " text frame"
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

// ===================== STYLES TAB (PHASE 1) =====================
function setupStylesTab(tab, defaultSettings, hasIndependentLayout, hasTableLayout, availableFields) {
    tab.alignChildren = 'fill';
    tab.margins = [15, 15, 15, 15];

    var paraStyles = getParagraphStyleNames();
    var tableStyles = getTableStyleNames();
    var cellStyles = getCellStyleNames();
    var items = ['— None —'];
    for (var i = 0; i < paraStyles.length; i++) items.push(paraStyles[i]);

    var itemsTable = ['— None —'];
    for (var ti = 0; ti < tableStyles.length; ti++) itemsTable.push(tableStyles[ti]);
    var itemsCell = ['— None —'];
    for (var ci = 0; ci < cellStyles.length; ci++) itemsCell.push(cellStyles[ci]);

    var pnlDesc = tab.add('panel', undefined, 'Styles (Phase 1: Paragraph Styles)');
    pnlDesc.margins = [15, 12, 15, 12];
    var desc = pnlDesc.add('statictext', undefined, 'Choose paragraph styles to apply. Table/cell styles will be added in Phase 2.');
    desc.maximumSize.width = 540;

    // Session fields
    var pnlSession = tab.add('panel', undefined, 'Session Fields');
    pnlSession.alignChildren = 'fill';
    pnlSession.margins = [15, 15, 15, 15];

    var g1 = pnlSession.add('group'); g1.orientation = 'row'; g1.spacing = 10;
    g1.add('statictext', undefined, 'Session Title:').preferredSize.width = 120;
    var ddSessionTitle = g1.add('dropdownlist', undefined, items); ddSessionTitle.selection = 0;

    var g2 = pnlSession.add('group'); g2.orientation = 'row'; g2.spacing = 10;
    g2.add('statictext', undefined, 'Session Time:').preferredSize.width = 120;
    var ddSessionTime = g2.add('dropdownlist', undefined, items); ddSessionTime.selection = 0;

    var g3 = pnlSession.add('group'); g3.orientation = 'row'; g3.spacing = 10;
    g3.add('statictext', undefined, 'Session No:').preferredSize.width = 120;
    var ddSessionNo = g3.add('dropdownlist', undefined, items); ddSessionNo.selection = 0;

    // Chairpersons
    var pnlChair = tab.add('panel', undefined, 'Chairpersons');
    pnlChair.alignChildren = 'fill';
    pnlChair.margins = [15, 15, 15, 15];

    var c1 = pnlChair.add('group'); c1.orientation = 'row'; c1.spacing = 10;
    c1.add('statictext', undefined, 'Inline mode:').preferredSize.width = 120;
    var ddChairInline = c1.add('dropdownlist', undefined, items); ddChairInline.selection = 0;

    var c2 = pnlChair.add('group'); c2.orientation = 'row'; c2.spacing = 10;
    c2.add('statictext', undefined, 'Grid mode:').preferredSize.width = 120;
    var ddChairGrid = c2.add('dropdownlist', undefined, items); ddChairGrid.selection = 0;

    // Topics (Table Layout)
    var pnlTopicsTable = tab.add('panel', undefined, 'Topics (Table Layout)');
    pnlTopicsTable.alignChildren = 'fill';
    pnlTopicsTable.margins = [15, 15, 15, 15];

    var tt1 = pnlTopicsTable.add('group'); tt1.orientation = 'row'; tt1.spacing = 10;
    tt1.add('statictext', undefined, 'Table Style:').preferredSize.width = 120;
    var ddTableStyle = tt1.add('dropdownlist', undefined, itemsTable); ddTableStyle.selection = 0;

    var tt2 = pnlTopicsTable.add('group'); tt2.orientation = 'row'; tt2.spacing = 10;
    tt2.add('statictext', undefined, 'Cell Style:').preferredSize.width = 120;
    var ddCellStyle = tt2.add('dropdownlist', undefined, itemsCell); ddCellStyle.selection = 0;

    // Topics (independent rows)
    var pnlTopicsInd = tab.add('panel', undefined, 'Topics (Independent Rows)');
    pnlTopicsInd.alignChildren = 'fill';
    pnlTopicsInd.margins = [15, 15, 15, 15];

    var t1 = pnlTopicsInd.add('group'); t1.orientation = 'row'; t1.spacing = 10;
    t1.add('statictext', undefined, 'Time:').preferredSize.width = 120;
    var ddTopicTime = t1.add('dropdownlist', undefined, items); ddTopicTime.selection = 0;

    var t2 = pnlTopicsInd.add('group'); t2.orientation = 'row'; t2.spacing = 10;
    t2.add('statictext', undefined, 'Title:').preferredSize.width = 120;
    var ddTopicTitle = t2.add('dropdownlist', undefined, items); ddTopicTitle.selection = 0;

    var t3 = pnlTopicsInd.add('group'); t3.orientation = 'row'; t3.spacing = 10;
    t3.add('statictext', undefined, 'Speaker:').preferredSize.width = 120;
    var ddTopicSpeaker = t3.add('dropdownlist', undefined, items); ddTopicSpeaker.selection = 0;

    // Enable/disable based on available fields and layout
    pnlTopicsInd.enabled = !!hasIndependentLayout;
    pnlTopicsTable.enabled = !!hasTableLayout;
    try {
        if (availableFields) {
            ddSessionTitle.enabled = !!availableFields['Session Title'];
            ddSessionTime.enabled = !!availableFields['Session Time'];
            ddSessionNo.enabled = !!availableFields['Session No'];
        }
    } catch (e) {}

    // Store refs for collection/import
    tab.ddSessionTitle = ddSessionTitle;
    tab.ddSessionTime = ddSessionTime;
    tab.ddSessionNo = ddSessionNo;
    tab.ddChairInline = ddChairInline;
    tab.ddChairGrid = ddChairGrid;
    tab.ddTableStyle = ddTableStyle;
    tab.ddCellStyle = ddCellStyle;
    tab.ddTopicTime = ddTopicTime;
    tab.ddTopicTitle = ddTopicTitle;
    tab.ddTopicSpeaker = ddTopicSpeaker;

    // Initialize from defaults
    var s = defaultSettings.stylesOptions || {};
    if (s.table) {
        selectDropdownByText(ddTableStyle, s.table.tableStyle);
        selectDropdownByText(ddCellStyle, s.table.cellStyle);
    }
    if (s.session) {
        selectDropdownByText(ddSessionTitle, s.session.titlePara);
        selectDropdownByText(ddSessionTime, s.session.timePara);
        selectDropdownByText(ddSessionNo, s.session.noPara);
    }
    if (s.chair) {
        selectDropdownByText(ddChairInline, s.chair.inlinePara);
        selectDropdownByText(ddChairGrid, s.chair.gridPara);
    }
    if (s.topicsIndependent) {
        selectDropdownByText(ddTopicTime, s.topicsIndependent.timePara);
        selectDropdownByText(ddTopicTitle, s.topicsIndependent.titlePara);
        selectDropdownByText(ddTopicSpeaker, s.topicsIndependent.speakerPara);
    }

    // ----- Create New Style (UI) -----
    var pnlCreate = tab.add('panel', undefined, 'Create New Style');
    pnlCreate.alignChildren = 'left';
    pnlCreate.margins = [15, 15, 15, 15];

    var grpRow1 = pnlCreate.add('group'); grpRow1.orientation = 'row'; grpRow1.spacing = 10;
    grpRow1.add('statictext', undefined, 'Name:').preferredSize.width = 60;
    var etStyleName = grpRow1.add('edittext', undefined, '');
    etStyleName.characters = 25;
    grpRow1.add('statictext', undefined, 'Type:').preferredSize.width = 40;
    var ddStyleType = grpRow1.add('dropdownlist', undefined, ['Paragraph', 'Table', 'Cell']);
    ddStyleType.selection = 0;
    var btnCreateStyle = grpRow1.add('button', undefined, 'Create');

    // Helper: refresh all dropdowns from current doc styles. If newName provided, auto-select where applicable
    function refreshStyleDropdowns(newName, newType) {
        try {
            // Rebuild paragraph items
            var pNames = getParagraphStyleNames();
            var pItems = ['— None —'];
            for (var i = 0; i < pNames.length; i++) pItems.push(pNames[i]);

            function repop(dd, itemsArr, targetName) {
                if (!dd) return;
                var prev = (dd.selection && dd.selection.text) ? dd.selection.text : '';
                dd.removeAll();
                for (var j = 0; j < itemsArr.length; j++) dd.add('item', itemsArr[j]);
                // Restore selection or select targetName
                var want = targetName || prev;
                selectDropdownByText(dd, want);
            }

            repop(ddSessionTitle, pItems, newType === 'Paragraph' ? newName : null);
            repop(ddSessionTime, pItems, null);
            repop(ddSessionNo, pItems, null);
            repop(ddChairInline, pItems, null);
            repop(ddChairGrid, pItems, null);
            repop(ddTopicTime, pItems, null);
            repop(ddTopicTitle, pItems, null);
            repop(ddTopicSpeaker, pItems, null);

            // Rebuild table style items
            var tNames = getTableStyleNames();
            var tItems = ['— None —'];
            for (var k = 0; k < tNames.length; k++) tItems.push(tNames[k]);
            repop(ddTableStyle, tItems, newType === 'Table' ? newName : null);

            // Rebuild cell style items
            var cNames = getCellStyleNames();
            var cItems = ['— None —'];
            for (var m = 0; m < cNames.length; m++) cItems.push(cNames[m]);
            repop(ddCellStyle, cItems, newType === 'Cell' ? newName : null);
        } catch (e) {}
    }

    // Create button handler
    btnCreateStyle.onClick = function() {
        var name = trimES(etStyleName.text);
        if (!name) { alert('Please enter a style name.'); return; }
        var type = (ddStyleType.selection && ddStyleType.selection.text) ? ddStyleType.selection.text : 'Paragraph';
        try {
            if (type === 'Paragraph') {
                createParagraphStyleIfMissing(name, { pointSize: 10, leading: 12 });
            } else if (type === 'Table') {
                // Try to reference body cell if exists
                var bodyCandidates = getCellStyleNames();
                var bodyName = bodyCandidates.length > 0 ? bodyCandidates[0] : undefined;
                var doc = app.activeDocument;
                var bodyCS = bodyName ? doc.cellStyles.itemByName(bodyName) : null;
                createTableStyleIfMissing(name, { bodyRegionCellStyle: (bodyCS && bodyCS.isValid) ? bodyCS : undefined });
            } else if (type === 'Cell') {
                createCellStyleIfMissing(name, { topInset: 1, bottomInset: 1, leftInset: 1, rightInset: 1 });
            }
            refreshStyleDropdowns(name, type);
            // Auto-select a sensible default dropdown for the created type
            if (type === 'Paragraph') {
                selectDropdownByText(ddSessionTitle, name);
            } else if (type === 'Table') {
                selectDropdownByText(ddTableStyle, name);
            } else if (type === 'Cell') {
                selectDropdownByText(ddCellStyle, name);
            }
            etStyleName.text = '';
        } catch (e) { alert('Failed to create style: ' + e); }
    };
}

function updateStylesTabFromSettings(tab, stylesOptions, hasIndependentLayout, hasTableLayout, availableFields) {
    if (!tab || !stylesOptions) return;
    try {
        if (stylesOptions.table) {
            selectDropdownByText(tab.ddTableStyle, stylesOptions.table.tableStyle);
            selectDropdownByText(tab.ddCellStyle, stylesOptions.table.cellStyle);
        }
        if (stylesOptions.session) {
            selectDropdownByText(tab.ddSessionTitle, stylesOptions.session.titlePara);
            selectDropdownByText(tab.ddSessionTime, stylesOptions.session.timePara);
            selectDropdownByText(tab.ddSessionNo, stylesOptions.session.noPara);
        }
        if (stylesOptions.chair) {
            selectDropdownByText(tab.ddChairInline, stylesOptions.chair.inlinePara);
            selectDropdownByText(tab.ddChairGrid, stylesOptions.chair.gridPara);
        }
        if (stylesOptions.topicsIndependent) {
            selectDropdownByText(tab.ddTopicTime, stylesOptions.topicsIndependent.timePara);
            selectDropdownByText(tab.ddTopicTitle, stylesOptions.topicsIndependent.titlePara);
            selectDropdownByText(tab.ddTopicSpeaker, stylesOptions.topicsIndependent.speakerPara);
        }
        if (tab.ddTopicTime) {
            var enableInd = !!hasIndependentLayout;
            tab.ddTopicTime.enabled = enableInd;
            tab.ddTopicTitle.enabled = enableInd;
            tab.ddTopicSpeaker.enabled = enableInd;
        }
        if (tab.ddTableStyle) {
            var enableTbl = !!hasTableLayout;
            tab.ddTableStyle.enabled = enableTbl;
            tab.ddCellStyle.enabled = enableTbl;
        }
        if (availableFields) {
            if (tab.ddSessionTitle) tab.ddSessionTitle.enabled = !!availableFields['Session Title'];
            if (tab.ddSessionTime) tab.ddSessionTime.enabled = !!availableFields['Session Time'];
            if (tab.ddSessionNo) tab.ddSessionNo.enabled = !!availableFields['Session No'];
        }
    } catch (e) {}
}

function getParagraphStyleNames() {
    var names = [];
    try {
        var doc = app.activeDocument;
        var arr = doc.allParagraphStyles ? doc.allParagraphStyles : doc.paragraphStyles;
        for (var i = 0; i < arr.length; i++) {
            try {
                var nm = arr[i].name;
                if (!nm) continue;
                if (nm.charAt(0) === '[') continue; // Skip [No Paragraph Style], [Basic Paragraph]
                if (indexOfExact(names, nm) === -1) names.push(nm);
            } catch (ignore) {}
        }
    } catch (e) {}
    names.sort();
    return names;
}

function selectDropdownByText(dd, text) {
    if (!dd) return;
    if (!text || text === '') { dd.selection = 0; return; }
    for (var i = 0; i < dd.items.length; i++) {
        if (dd.items[i].text === text) { dd.selection = i; return; }
    }
    dd.selection = 0;
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

/* ---------------- MASTER (PARENT) SPREAD HANDLING ---------------- */

// Copy all master spreads from source to target document.
// Returns a map: { masterName: MasterSpread-in-target }
function copyMasterSpreads(sourceDoc, targetDoc) {
    var map = {};
    try {
        var srcMasters = sourceDoc.masterSpreads;
        for (var i = 0; i < srcMasters.length; i++) {
            var src = srcMasters[i];
            var name = "";
            try { name = src.name; } catch (eName) { name = "" + i; }

            var duplicated = null;
            try {
                // Primary attempt: duplicate master into target document
                duplicated = src.duplicate(LocationOptions.AT_END, targetDoc);
            } catch (e1) {
                $.writeln("Master duplicate failed for '" + name + "' -> " + e1);
            }

            if (!duplicated) {
                try {
                    // Fallback: if duplication signature differs, try duplicating after last master
                    duplicated = src.duplicate(LocationOptions.AFTER, targetDoc.masterSpreads.lastItem());
                } catch (e2) {
                    $.writeln("Master duplicate (fallback) failed for '" + name + "' -> " + e2);
                }
            }

            if (!duplicated) {
                // Final fallback: try to reference by name if already present (avoid duplicates)
                try {
                    duplicated = targetDoc.masterSpreads.itemByName(name);
                    if (!duplicated.isValid) duplicated = null;
                } catch (e3) { duplicated = null; }
            }

            if (duplicated) {
                map[name] = duplicated;
            }
        }
        $.writeln("Copied " + (function(){var c=0; for (var k in map) c++; return c;})() + " master spreads to target document.");
    } catch (e) {
        $.writeln("Warning: Failed while copying master spreads: " + e);
    }
    return map;
}

// Apply masters to targetDoc pages based on the corresponding template pages.
// Assumes the first N pages in targetDoc correspond to sourceDoc.pages after initial cleanup.
function applyMastersToPages(sourceDoc, targetDoc, masterMap) {
    try {
        var n = Math.min(sourceDoc.pages.length, targetDoc.pages.length);
        for (var i = 0; i < n; i++) {
            var srcPage = sourceDoc.pages[i];
            var tgtPage = targetDoc.pages[i];
            try {
                var srcMaster = srcPage.appliedMaster;
                if (srcMaster) {
                    var name = "";
                    try { name = srcMaster.name; } catch (eName) { name = ""; }
                    var tgtMaster = masterMap && masterMap[name] ? masterMap[name] : null;
                    if (!tgtMaster) {
                        // Try to find by name on demand
                        try {
                            var candidate = targetDoc.masterSpreads.itemByName(name);
                            if (candidate && candidate.isValid) tgtMaster = candidate;
                        } catch (eFind) {}
                    }
                    if (tgtMaster) {
                        tgtPage.appliedMaster = tgtMaster;
                    }
                }
            } catch (inner) {
                // Continue with next page
            }
        }
        $.writeln("Applied masters to initial duplicated pages.");
    } catch (e) {
        $.writeln("Warning: Could not apply masters to pages: " + e);
    }
}