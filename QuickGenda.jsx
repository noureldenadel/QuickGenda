// QuickGenda (v2.0 - Styling and Layout)
// Date: August 22, 2025
// Author: noorr
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
function setupFormattingCombinedTab(tab, defaultSettings, hasIndependentLayout, hasTableLayout, availableFields, detectedTemplateStyles) {
    tab.alignChildren = 'fill';
    tab.margins = [16, 16, 16, 16]; // 4px spacing system compliance

    function getParagraphStyleNames() {
        var names = ['— None —'];
        try {
            var ps = app.activeDocument.paragraphStyles;
            for (var i = 0; i < ps.length; i++) {
                var nm = ps[i].name;
                if (nm && nm.charAt(0) !== '[') names.push(nm);
            }
        } catch (e) {}
        return names;
    }
    function getTableStyleNames() {
        var names = ['— None —'];
        try { 
            var ts = app.activeDocument.tableStyles; 
            for (var i=0;i<ts.length;i++) {
                var nm = ts[i].name;
                if (nm && nm.charAt(0) !== '[') names.push(nm);
            }
        } catch(e) {}
        return names;
    }
    function getCellStyleNames() {
        var names = ['— None —'];
        try { 
            var cs = app.activeDocument.cellStyles; 
            for (var i=0;i<cs.length;i++) {
                var nm = cs[i].name;
                if (nm && nm.charAt(0) !== '[') names.push(nm);
            }
        } catch(e) {}
        return names;
    }

    function addStyleRow(parent, labelText, ddItems, preselectText, defaultEnabled, defaultChar, isTableOrCell) {
        var grp = parent.add('group'); 
        grp.orientation = 'row'; 
        grp.spacing = 8; // 4px system: medium spacing
        grp.alignChildren = 'left';
        
        var st = grp.add('statictext', undefined, labelText); 
        st.preferredSize.width = 120; // 4px system: medium label width
        
        var dd = grp.add('dropdownlist', undefined, ddItems); 
        dd.preferredSize.width = 200;
        
        // Priority-based selection: detected style > "— None —"
        var selectionMade = false;
        if (preselectText && preselectText !== null && preselectText !== '') {
            for (var i = 0; i < dd.items.length; i++) {
                if (dd.items[i].text === preselectText) {
                    dd.selection = i;
                    selectionMade = true;
                    break;
                }
            }
        }
        if (!selectionMade) {
            dd.selection = 0;
        }
        
        // Add line break controls only for text fields (not table/cell styles)
        var breakControls = null;
        if (!isTableOrCell) {
            grp.add('statictext', undefined, '    '); // 4px spacing spacer
            var cb = grp.add('checkbox', undefined, ''); 
            cb.value = !!defaultEnabled;
            cb.enabled = true; // Always enabled so user can control the other elements
            
            var et = grp.add('edittext', undefined, (defaultChar || '|')); 
            et.characters = 2; // Half size as requested
            et.enabled = !!defaultEnabled; // Enable if default is true
            
            var arrow = grp.add('statictext', undefined, '-> ¶');
            arrow.enabled = !!defaultEnabled; // Enable if default is true
            
            // Enable/disable controls based on checkbox state
            cb.onClick = function() {
                et.enabled = cb.value;
                arrow.enabled = cb.value;
                if (cb.value) {
                    et.active = true;
                }
            };
            
            breakControls = { checkbox: cb, charInput: et, arrow: arrow };
        }
        
        return { 
            group: grp, 
            label: st, 
            dropdown: dd,
            breakControls: breakControls
        };
    }

    // Use the detected styles from the original template document
    var detectedStyles = detectedTemplateStyles || {
        sessionTitle: null,
        sessionTime: null,
        sessionNo: null,
        chairStyle: null,
        topicTime: null,
        topicTitle: null,
        topicSpeaker: null,
        tableStyle: null,
        cellStyle: null
    };
    
    // Single Fields Panel containing all field types
    var pnlFields = tab.add('panel', undefined, 'Fields');
    pnlFields.margins = [16, 16, 16, 16]; // 4px system compliance
    pnlFields.alignChildren = 'left';
    
    var paraItems = getParagraphStyleNames();
    var tableItems = getTableStyleNames();
    var cellItems = getCellStyleNames();

    // Session Fields
    var rowTitle = addStyleRow(pnlFields, 'Session Title:', paraItems, 
        (defaultSettings.stylesOptions && defaultSettings.stylesOptions.session && defaultSettings.stylesOptions.session.titlePara) ? defaultSettings.stylesOptions.session.titlePara : detectedStyles.sessionTitle,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionTitle'] ? defaultSettings.lineBreaks['SessionTitle'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionTitle'] ? defaultSettings.lineBreaks['SessionTitle'].character : '|');

    var rowTime = addStyleRow(pnlFields, 'Session Time:', paraItems, 
        (defaultSettings.stylesOptions && defaultSettings.stylesOptions.session && defaultSettings.stylesOptions.session.timePara) ? defaultSettings.stylesOptions.session.timePara : detectedStyles.sessionTime,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionTime'] ? defaultSettings.lineBreaks['SessionTime'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionTime'] ? defaultSettings.lineBreaks['SessionTime'].character : '|');

    var rowNo = addStyleRow(pnlFields, 'Session Number:', paraItems, 
        (defaultSettings.stylesOptions && defaultSettings.stylesOptions.session && defaultSettings.stylesOptions.session.noPara) ? defaultSettings.stylesOptions.session.noPara : detectedStyles.sessionNo,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionNo'] ? defaultSettings.lineBreaks['SessionNo'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['SessionNo'] ? defaultSettings.lineBreaks['SessionNo'].character : '|');

    // Chairpersons
    var rowChair = addStyleRow(pnlFields, 'Chairpersons:', paraItems, 
        (defaultSettings.stylesOptions && defaultSettings.stylesOptions.chair && defaultSettings.stylesOptions.chair.style) ? defaultSettings.stylesOptions.chair.style : detectedStyles.chairStyle,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['Chairpersons'] ? defaultSettings.lineBreaks['Chairpersons'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['Chairpersons'] ? defaultSettings.lineBreaks['Chairpersons'].character : '||');

    // Topic Fields
    var rowTopicTime = addStyleRow(pnlFields, 'Topic Time:', paraItems, 
        (defaultSettings.stylesOptions && defaultSettings.stylesOptions.topicsIndependent && defaultSettings.stylesOptions.topicsIndependent.timePara) ? defaultSettings.stylesOptions.topicsIndependent.timePara : detectedStyles.topicTime,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicTime'] ? defaultSettings.lineBreaks['topicTime'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicTime'] ? defaultSettings.lineBreaks['topicTime'].character : '|');

    var rowTopicTitle = addStyleRow(pnlFields, 'Topic Title:', paraItems, 
        (defaultSettings.stylesOptions && defaultSettings.stylesOptions.topicsIndependent && defaultSettings.stylesOptions.topicsIndependent.titlePara) ? defaultSettings.stylesOptions.topicsIndependent.titlePara : detectedStyles.topicTitle,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicTitle'] ? defaultSettings.lineBreaks['topicTitle'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicTitle'] ? defaultSettings.lineBreaks['topicTitle'].character : '|');

    var rowTopicSpeaker = addStyleRow(pnlFields, 'Topic Speaker:', paraItems, 
        (defaultSettings.stylesOptions && defaultSettings.stylesOptions.topicsIndependent && defaultSettings.stylesOptions.topicsIndependent.speakerPara) ? defaultSettings.stylesOptions.topicsIndependent.speakerPara : detectedStyles.topicSpeaker,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicSpeaker'] ? defaultSettings.lineBreaks['topicSpeaker'].enabled : false,
        defaultSettings.lineBreaks && defaultSettings.lineBreaks['topicSpeaker'] ? defaultSettings.lineBreaks['topicSpeaker'].character : '|');

    // Table and Cell Styles (no line break controls)
    var rowTbl = addStyleRow(pnlFields, 'Table Style:', tableItems, 
        (defaultSettings.stylesOptions && defaultSettings.stylesOptions.table && defaultSettings.stylesOptions.table.tableStyle) ? defaultSettings.stylesOptions.table.tableStyle : detectedStyles.tableStyle,
        false, '', true);

    var rowCell = addStyleRow(pnlFields, 'Cell Style:', cellItems, 
        (defaultSettings.stylesOptions && defaultSettings.stylesOptions.table && defaultSettings.stylesOptions.table.cellStyle) ? defaultSettings.stylesOptions.table.cellStyle : detectedStyles.cellStyle,
        false, '', true);

    // Create New Style Panel
    var pnlNewStyle = tab.add('panel', undefined, 'Create New Style');
    pnlNewStyle.alignChildren = 'left';
    pnlNewStyle.margins = [16, 16, 16, 16]; // 4px system compliance
    
    var grpNewStyle = pnlNewStyle.add('group'); 
    grpNewStyle.orientation = 'row'; 
    grpNewStyle.spacing = 8; // 4px system: medium spacing
    grpNewStyle.alignChildren = 'left';
    
    grpNewStyle.add('statictext', undefined, 'Name:').preferredSize.width = 50;
    var etNewStyleName = grpNewStyle.add('edittext', undefined, ''); 
    etNewStyleName.characters = 20;
    
    grpNewStyle.add('statictext', undefined, 'Type:').preferredSize.width = 35;
    var ddNewStyleType = grpNewStyle.add('dropdownlist', undefined, ['Paragraph', 'Table', 'Cell']); 
    ddNewStyleType.selection = 0;
    ddNewStyleType.preferredSize.width = 100;
    
    var btnCreate = grpNewStyle.add('button', undefined, '+ Create');
    btnCreate.preferredSize.width = 80;
    
    btnCreate.onClick = function() {
        var name = etNewStyleName.text;
        if (!name || name === '') {
            alert('Please enter a style name.');
            return;
        }
        var type = ddNewStyleType.selection.text;
        try {
            if (type === 'Paragraph') { 
                app.activeDocument.paragraphStyles.add({name: name}); 
            } else if (type === 'Table') { 
                app.activeDocument.tableStyles.add({name: name}); 
            } else { 
                app.activeDocument.cellStyles.add({name: name}); 
            }
            etNewStyleName.text = '';
            
            // repopulate dropdowns
            var pItems = getParagraphStyleNames(); var tItems = getTableStyleNames(); var cItems = getCellStyleNames();
            function repop(dd, items){ 
                if (!dd) return; 
                var sel = (dd.selection && dd.selection.text) ? dd.selection.text : null; 
                dd.removeAll(); 
                for (var i=0;i<items.length;i++) dd.add('item', items[i]); 
                dd.selection = sel ? (function(){ 
                    for (var j=0;j<dd.items.length;j++){ 
                        if(dd.items[j].text===sel) return j;
                    } 
                    return 0;
                })() : 0; 
            }
            repop(rowTitle.dropdown, pItems); repop(rowTime.dropdown, pItems); repop(rowNo.dropdown, pItems);
            repop(rowChair.dropdown, pItems);  // Updated to use single chair dropdown
            repop(rowTopicTime.dropdown, pItems); repop(rowTopicTitle.dropdown, pItems); repop(rowTopicSpeaker.dropdown, pItems);
            repop(rowTbl.dropdown, tItems); repop(rowCell.dropdown, cItems);
            
            // Don't auto-select the newly created style in any dropdown
            // User should manually select the style they want to apply
        } catch (e) { 
            alert('Failed to create style: ' + e); 
        }
    };

    // Smart hiding of elements based on CSV availability and layout capabilities
    // Helper functions for proper gap-free layout management
    function forceRelayout(ctrl) {
        try {
            var root = ctrl;
            // climb to top-most parent (Window)
            while (root && root.parent) root = root.parent;
            if (root && root.layout) {
                root.layout.layout(true);
                root.layout.resize();
                // extra layout pass to stabilize
                root.layout.layout(true);
            }
            if (typeof root.update === 'function') root.update();
        } catch (e) {}
    }
    function ensureCollapsible(panel) { 
        try { if (panel && panel.minimumSize) panel.minimumSize.height = 0; } catch (e) {} 
    }
    function setCollapsed(panel, collapsed) {
        try {
            if (!panel) return;
            if (collapsed) {
                if (panel.minimumSize) panel.minimumSize.height = 0;
                if (panel.maximumSize) panel.maximumSize.height = 0;
                panel.visible = false;
            } else {
                if (panel.maximumSize) panel.maximumSize.height = 10000;
                panel.visible = true;
            }
        } catch (e) {}
    }
    
    // Prepare all elements for potential collapsing
    ensureCollapsible(rowTitle.group);
    ensureCollapsible(rowTime.group);
    ensureCollapsible(rowNo.group);
    ensureCollapsible(rowChair.group);
    ensureCollapsible(rowTopicTime.group);
    ensureCollapsible(rowTopicTitle.group);
    ensureCollapsible(rowTopicSpeaker.group);
    ensureCollapsible(rowTbl.group);
    ensureCollapsible(rowCell.group);
    ensureCollapsible(pnlFields);
    
    try {
        if (availableFields) {
            // Hide individual session fields that aren't in CSV
            if (!availableFields['Session Title']) { 
                setCollapsed(rowTitle.group, true); 
            }
            if (!availableFields['Session Time']) { 
                setCollapsed(rowTime.group, true); 
            }
            if (!availableFields['Session No']) { 
                setCollapsed(rowNo.group, true); 
            }
            
            // Hide Chairpersons if not in CSV
            if (!availableFields['Chairpersons']) { setCollapsed(rowChair.group, true); }
            
            // Hide individual topic fields that aren't in CSV
            if (!availableFields['Time']) { 
                setCollapsed(rowTopicTime.group, true); 
            }
            if (!availableFields['Topic Title']) { 
                setCollapsed(rowTopicTitle.group, true); 
            }
            if (!availableFields['Speaker']) { 
                setCollapsed(rowTopicSpeaker.group, true); 
            }
        }
        
        // Hide table/cell styles based on template layout capabilities
        if (!hasTableLayout) { 
            setCollapsed(rowTbl.group, true); 
            setCollapsed(rowCell.group, true); 
        }
        
        // Hide independent topic fields if layout not supported
        if (!hasIndependentLayout) { 
            setCollapsed(rowTopicTime.group, true);
            setCollapsed(rowTopicTitle.group, true);
            setCollapsed(rowTopicSpeaker.group, true);
        }
        
        // Force comprehensive layout update to remove all gaps
        forceRelayout(tab);
        
    } catch (e) {}

    // Expose properties so existing collectors work
    tab.ddSessionTitle = rowTitle.dropdown;
    tab.ddSessionTime = rowTime.dropdown;
    tab.ddSessionNo = rowNo.dropdown;
    tab.ddChair = rowChair.dropdown;
    tab.ddTableStyle = rowTbl.dropdown;
    tab.ddCellStyle = rowCell.dropdown;
    tab.ddTopicTime = rowTopicTime.dropdown;
    tab.ddTopicTitle = rowTopicTitle.dropdown;
    tab.ddTopicSpeaker = rowTopicSpeaker.dropdown;

    tab.sessionControls = {
        'SessionTitle': { checkbox: rowTitle.breakControls.checkbox, charInput: rowTitle.breakControls.charInput },
        'SessionTime': { checkbox: rowTime.breakControls.checkbox, charInput: rowTime.breakControls.charInput },
        'SessionNo': { checkbox: rowNo.breakControls.checkbox, charInput: rowNo.breakControls.charInput }
    };
    tab.chairpersonControls = {
        'Chairpersons': { checkbox: rowChair.breakControls.checkbox, charInput: rowChair.breakControls.charInput }
    };
    tab.topicControls = {
        'topicTime': { checkbox: rowTopicTime.breakControls.checkbox, charInput: rowTopicTime.breakControls.charInput },
        'topicTitle': { checkbox: rowTopicTitle.breakControls.checkbox, charInput: rowTopicTitle.breakControls.charInput },
        'topicSpeaker': { checkbox: rowTopicSpeaker.breakControls.checkbox, charInput: rowTopicSpeaker.breakControls.charInput }
    };}

function main() {
    if (app.documents.length === 0) {
        alert("Error: No document is open.\nPlease open your template file first.");
        return;
    }

    var templateDoc = app.activeDocument;

    // CRITICAL: Detect styles from original template BEFORE creating new document
    var detectedTemplateStyles = detectActiveStylesFromTemplateDoc(templateDoc);

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
    
    // CRITICAL: Re-verify and re-apply document setup after all page operations
    // Some InDesign operations (master copying, page duplication) can override document settings
    
    // If settings changed, reapply them
    var sourceFacingPages = templateDoc.documentPreferences.facingPages;
    var sourcePageBinding = templateDoc.documentPreferences.pageBinding;
    var sourcePageDirection = safeGetDocProperty(templateDoc, 'pageDirection', 'N/A');
    
    var currentFacingPages = doc.documentPreferences.facingPages;
    var currentPageBinding = doc.documentPreferences.pageBinding;
    var currentPageDirection = safeGetDocProperty(doc, 'pageDirection', 'N/A');
    
    var facingPagesChanged = currentFacingPages !== sourceFacingPages;
    var pageBindingChanged = currentPageBinding !== sourcePageBinding;
    var pageDirectionChanged = (sourcePageDirection !== 'N/A') && (currentPageDirection !== sourcePageDirection);
    
    if (facingPagesChanged || pageBindingChanged || pageDirectionChanged) {
        // Re-apply critical document setup settings
        doc.documentPreferences.facingPages = sourceFacingPages;
        doc.documentPreferences.pageBinding = sourcePageBinding;
        
        if (sourcePageDirection !== 'N/A') {
            safeSetDocProperty(doc, 'pageDirection', sourcePageDirection);
        }
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

    // Get all user options from unified settings panel (pass detected template styles)
    var allOptions = getUnifiedSettingsPanel(hasIndependentTopicLayout, hasTableTopicLayout, availableFields, csvInfo, detectedTemplateStyles);
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
        var reportFile = exportReport(sessionsData, allOptions.reportOptions);
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
// ensureDefaultStyles function removed - auto-style creation feature disabled
// Manual style creation is still available through the UI


/* ---------------- LAYOUT DETECTION ---------------- */

// Function to detect active styles from a specific template document
function detectActiveStylesFromTemplateDoc(templateDoc) {
    var activeStyles = {
        sessionTitle: null,
        sessionTime: null,
        sessionNo: null,
        chairStyle: null,
        topicTime: null,
        topicTitle: null,
        topicSpeaker: null,
        tableStyle: null,
        cellStyle: null
    };
    
    try {
        var templatePage = templateDoc.pages[0];
        
        // Detect paragraph styles from text frames
        var sessionTitleFrame = findItemByLabel(templatePage, 'sessionTitle');
        if (sessionTitleFrame) {
            try {
                if (sessionTitleFrame.constructor.name === "TextFrame") {
                    var styleName = null;
                    
                    // Method 1: Direct access to appliedParagraphStyle
                    try {
                        if (sessionTitleFrame.appliedParagraphStyle && sessionTitleFrame.appliedParagraphStyle.name) {
                            styleName = sessionTitleFrame.appliedParagraphStyle.name;
                        }
                    } catch (e1) {}
                    
                    // Method 2: Access via text content if direct access fails
                    if (!styleName) {
                        try {
                            if (sessionTitleFrame.contents && sessionTitleFrame.contents.length > 0) {
                                if (sessionTitleFrame.parentStory && sessionTitleFrame.parentStory.paragraphs.length > 0) {
                                    var firstPara = sessionTitleFrame.parentStory.paragraphs[0];
                                    if (firstPara.appliedParagraphStyle && firstPara.appliedParagraphStyle.name) {
                                        styleName = firstPara.appliedParagraphStyle.name;
                                    }
                                }
                            }
                        } catch (e2) {}
                    }
                    
                    // Method 3: Access via text selection if other methods fail
                    if (!styleName) {
                        try {
                            if (sessionTitleFrame.texts.length > 0) {
                                var firstText = sessionTitleFrame.texts[0];
                                if (firstText.appliedParagraphStyle && firstText.appliedParagraphStyle.name) {
                                    styleName = firstText.appliedParagraphStyle.name;
                                }
                            }
                        } catch (e3) {}
                    }
                    
                    // Process the detected style
                    if (styleName && styleName !== '[Basic Paragraph]' && styleName.charAt(0) !== '[') {
                        activeStyles.sessionTitle = styleName;
                    }
                }
            } catch (e) {}
        }
        
        var sessionTimeFrame = findItemByLabel(templatePage, 'sessionTime');
        if (sessionTimeFrame) {
            try {
                if (sessionTimeFrame.constructor.name === "TextFrame") {
                    var styleName = null;
                    
                    // Method 1: Direct access
                    try {
                        if (sessionTimeFrame.appliedParagraphStyle && sessionTimeFrame.appliedParagraphStyle.name) {
                            styleName = sessionTimeFrame.appliedParagraphStyle.name;
                        }
                    } catch (e1) {}
                    
                    // Method 2: Via paragraph
                    if (!styleName) {
                        try {
                            if (sessionTimeFrame.contents && sessionTimeFrame.parentStory && sessionTimeFrame.parentStory.paragraphs.length > 0) {
                                var firstPara = sessionTimeFrame.parentStory.paragraphs[0];
                                if (firstPara.appliedParagraphStyle && firstPara.appliedParagraphStyle.name) {
                                    styleName = firstPara.appliedParagraphStyle.name;
                                }
                            }
                        } catch (e2) {}
                    }
                    
                    // Method 3: Via text
                    if (!styleName) {
                        try {
                            if (sessionTimeFrame.texts.length > 0) {
                                var firstText = sessionTimeFrame.texts[0];
                                if (firstText.appliedParagraphStyle && firstText.appliedParagraphStyle.name) {
                                    styleName = firstText.appliedParagraphStyle.name;
                                }
                            }
                        } catch (e3) {}
                    }
                    
                    if (styleName && styleName !== '[Basic Paragraph]' && styleName.charAt(0) !== '[') {
                        activeStyles.sessionTime = styleName;
                    }
                }
            } catch (e) {}
        }
        
        var sessionNoFrame = findItemByLabel(templatePage, 'sessionNo');
        if (sessionNoFrame) {
            try {
                if (sessionNoFrame.constructor.name === "TextFrame") {
                    if (sessionNoFrame.appliedParagraphStyle) {
                        var styleName = sessionNoFrame.appliedParagraphStyle.name;
                        if (styleName && styleName !== '[Basic Paragraph]' && styleName.charAt(0) !== '[') {
                            activeStyles.sessionNo = styleName;
                        }
                    }
                }
            } catch (e) {}
        }
        
        // For chairpersons, use a single style for both inline and grid modes
        var chairFrame = findItemByLabel(templatePage, 'chairpersons');
        if (chairFrame) {
            try {
                if (chairFrame.constructor.name === "TextFrame") {
                    var styleName = null;
                    
                    // Method 1: Direct access
                    try {
                        if (chairFrame.appliedParagraphStyle && chairFrame.appliedParagraphStyle.name) {
                            styleName = chairFrame.appliedParagraphStyle.name;
                        }
                    } catch (e1) {}
                    
                    // Method 2: Via paragraph
                    if (!styleName) {
                        try {
                            if (chairFrame.contents && chairFrame.parentStory && chairFrame.parentStory.paragraphs.length > 0) {
                                var firstPara = chairFrame.parentStory.paragraphs[0];
                                if (firstPara.appliedParagraphStyle && firstPara.appliedParagraphStyle.name) {
                                    styleName = firstPara.appliedParagraphStyle.name;
                                }
                            }
                        } catch (e2) {}
                    }
                    
                    // Method 3: Via text
                    if (!styleName) {
                        try {
                            if (chairFrame.texts.length > 0) {
                                var firstText = chairFrame.texts[0];
                                if (firstText.appliedParagraphStyle && firstText.appliedParagraphStyle.name) {
                                    styleName = firstText.appliedParagraphStyle.name;
                                }
                            }
                        } catch (e3) {}
                    }
                    
                    if (styleName && styleName !== '[Basic Paragraph]' && styleName.charAt(0) !== '[') {
                        activeStyles.chairStyle = styleName;
                    }
                }
            } catch (e) {}
        }
        
        // Detect independent topic styles (check both page level and inside groups)
        var topicTimeFrame = findItemByLabel(templatePage, 'topicTime');
        if (!topicTimeFrame) {
            // Search inside all groups on the page
            var allItems = templatePage.allPageItems;
            for (var i = 0; i < allItems.length && !topicTimeFrame; i++) {
                try {
                    if (allItems[i].constructor.name === "Group") {
                        topicTimeFrame = findDescendantByLabel(allItems[i], 'topicTime');
                    }
                } catch (e) {}
            }
        }
        if (topicTimeFrame) {
            try {
                if (topicTimeFrame.constructor.name === "TextFrame") {
                    var styleName = null;
                    
                    // Method 1: Direct access
                    try {
                        if (topicTimeFrame.appliedParagraphStyle && topicTimeFrame.appliedParagraphStyle.name) {
                            styleName = topicTimeFrame.appliedParagraphStyle.name;
                        }
                    } catch (e1) {}
                    
                    // Method 2: Via paragraph
                    if (!styleName) {
                        try {
                            if (topicTimeFrame.contents && topicTimeFrame.parentStory && topicTimeFrame.parentStory.paragraphs.length > 0) {
                                var firstPara = topicTimeFrame.parentStory.paragraphs[0];
                                if (firstPara.appliedParagraphStyle && firstPara.appliedParagraphStyle.name) {
                                    styleName = firstPara.appliedParagraphStyle.name;
                                }
                            }
                        } catch (e2) {}
                    }
                    
                    // Method 3: Via text
                    if (!styleName) {
                        try {
                            if (topicTimeFrame.texts.length > 0) {
                                var firstText = topicTimeFrame.texts[0];
                                if (firstText.appliedParagraphStyle && firstText.appliedParagraphStyle.name) {
                                    styleName = firstText.appliedParagraphStyle.name;
                                }
                            }
                        } catch (e3) {}
                    }
                    
                    if (styleName && styleName !== '[Basic Paragraph]' && styleName.charAt(0) !== '[') {
                        activeStyles.topicTime = styleName;
                    }
                }
            } catch (e) {}
        }
        
        var topicTitleFrame = findItemByLabel(templatePage, 'topicTitle');
        if (!topicTitleFrame) {
            // Search inside all groups on the page
            var allItems = templatePage.allPageItems;
            for (var i = 0; i < allItems.length && !topicTitleFrame; i++) {
                try {
                    if (allItems[i].constructor.name === "Group") {
                        topicTitleFrame = findDescendantByLabel(allItems[i], 'topicTitle');
                    }
                } catch (e) {}
            }
        }
        if (topicTitleFrame) {
            try {
                if (topicTitleFrame.constructor.name === "TextFrame") {
                    var styleName = null;
                    
                    // Method 1: Direct access
                    try {
                        if (topicTitleFrame.appliedParagraphStyle && topicTitleFrame.appliedParagraphStyle.name) {
                            styleName = topicTitleFrame.appliedParagraphStyle.name;
                        }
                    } catch (e1) {}
                    
                    // Method 2: Via paragraph
                    if (!styleName) {
                        try {
                            if (topicTitleFrame.contents && topicTitleFrame.parentStory && topicTitleFrame.parentStory.paragraphs.length > 0) {
                                var firstPara = topicTitleFrame.parentStory.paragraphs[0];
                                if (firstPara.appliedParagraphStyle && firstPara.appliedParagraphStyle.name) {
                                    styleName = firstPara.appliedParagraphStyle.name;
                                }
                            }
                        } catch (e2) {}
                    }
                    
                    // Method 3: Via text
                    if (!styleName) {
                        try {
                            if (topicTitleFrame.texts.length > 0) {
                                var firstText = topicTitleFrame.texts[0];
                                if (firstText.appliedParagraphStyle && firstText.appliedParagraphStyle.name) {
                                    styleName = firstText.appliedParagraphStyle.name;
                                }
                            }
                        } catch (e3) {}
                    }
                    
                    if (styleName && styleName !== '[Basic Paragraph]' && styleName.charAt(0) !== '[') {
                        activeStyles.topicTitle = styleName;
                    }
                }
            } catch (e) {}
        }
        
        var topicSpeakerFrame = findItemByLabel(templatePage, 'topicSpeaker');
        if (!topicSpeakerFrame) {
            // Search inside all groups on the page
            var allItems = templatePage.allPageItems;
            for (var i = 0; i < allItems.length && !topicSpeakerFrame; i++) {
                try {
                    if (allItems[i].constructor.name === "Group") {
                        topicSpeakerFrame = findDescendantByLabel(allItems[i], 'topicSpeaker');
                    }
                } catch (e) {}
            }
        }
        if (topicSpeakerFrame) {
            try {
                if (topicSpeakerFrame.constructor.name === "TextFrame") {
                    var styleName = null;
                    
                    // Method 1: Direct access
                    try {
                        if (topicSpeakerFrame.appliedParagraphStyle && topicSpeakerFrame.appliedParagraphStyle.name) {
                            styleName = topicSpeakerFrame.appliedParagraphStyle.name;
                        }
                    } catch (e1) {}
                    
                    // Method 2: Via paragraph
                    if (!styleName) {
                        try {
                            if (topicSpeakerFrame.contents && topicSpeakerFrame.parentStory && topicSpeakerFrame.parentStory.paragraphs.length > 0) {
                                var firstPara = topicSpeakerFrame.parentStory.paragraphs[0];
                                if (firstPara.appliedParagraphStyle && firstPara.appliedParagraphStyle.name) {
                                    styleName = firstPara.appliedParagraphStyle.name;
                                }
                            }
                        } catch (e2) {}
                    }
                    
                    // Method 3: Via text
                    if (!styleName) {
                        try {
                            if (topicSpeakerFrame.texts.length > 0) {
                                var firstText = topicSpeakerFrame.texts[0];
                                if (firstText.appliedParagraphStyle && firstText.appliedParagraphStyle.name) {
                                    styleName = firstText.appliedParagraphStyle.name;
                                }
                            }
                        } catch (e3) {}
                    }
                    
                    if (styleName && styleName !== '[Basic Paragraph]' && styleName.charAt(0) !== '[') {
                        activeStyles.topicSpeaker = styleName;
                    }
                }
            } catch (e) {}
        }
        
        // Detect table styles from topicsTable
        var topicsTableFrame = findItemByLabel(templatePage, 'topicsTable');
        if (topicsTableFrame && topicsTableFrame.hasOwnProperty('tables') && topicsTableFrame.tables.length > 0) {
            var table = topicsTableFrame.tables[0];
            
            // Get table style
            if (table.appliedTableStyle && table.appliedTableStyle.name) {
                var tableStyleName = table.appliedTableStyle.name;
                if (tableStyleName && tableStyleName !== '[Basic Table]' && tableStyleName.charAt(0) !== '[') {
                    activeStyles.tableStyle = tableStyleName;
                }
            }
            
            // Get cell style from first cell
            if (table.cells.length > 0 && table.cells[0].appliedCellStyle && table.cells[0].appliedCellStyle.name) {
                var cellStyleName = table.cells[0].appliedCellStyle.name;
                if (cellStyleName && cellStyleName !== '[None]' && cellStyleName.charAt(0) !== '[') {
                    activeStyles.cellStyle = cellStyleName;
                }
            }
        }
        
    } catch (e) {
        // Silently handle errors - if we can't detect styles, we'll just use defaults
        $.writeln("Warning: Could not detect template styles: " + e);
    }
    
    return activeStyles;
}

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

function getUnifiedSettingsPanel(hasIndependentTopicLayout, hasTableTopicLayout, availableFields, csvInfo, detectedTemplateStyles) {
    var defaultSettings = getDefaultSettings();
    // Auto-style creation feature removed - styles should be manually created or detected from template
    if (!defaultSettings.stylesOptions) defaultSettings.stylesOptions = {};
    if (!defaultSettings.stylesOptions.session) defaultSettings.stylesOptions.session = {};
    if (!defaultSettings.stylesOptions.table) defaultSettings.stylesOptions.table = {};
    if (!defaultSettings.stylesOptions.chair) defaultSettings.stylesOptions.chair = defaultSettings.stylesOptions.chair || {};
    if (!defaultSettings.stylesOptions.topicsIndependent) defaultSettings.stylesOptions.topicsIndependent = defaultSettings.stylesOptions.topicsIndependent || {};
    var dlg = new Window('dialog', 'QuickGenda - Configuration');
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
    // Ensure full-width, left-aligned sections for Advanced tab
    try { advancedTab.alignChildren = 'fill'; advancedTab.margins = [15, 15, 15, 15]; } catch (e) {}
    
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
        availableFields,
        detectedTemplateStyles
    );
    
    // ===== ADVANCED: CENTRALIZED SETTINGS MANAGEMENT =====
    // Top strip: Import/Export all (entire configuration)
    var pnlSettings = advancedTab.add('panel', undefined, 'Settings Management');
    pnlSettings.orientation = 'row';
    pnlSettings.alignChildren = 'left';
    pnlSettings.alignment = 'fill';
    pnlSettings.margins = [15, 15, 15, 15];
    var btnImport = pnlSettings.add('button', undefined, 'Import All Settings');
    var btnExport = pnlSettings.add('button', undefined, 'Export All Settings');
    btnImport.preferredSize.width = 150;
    btnExport.preferredSize.width = 150;
    

    // Report preferences
    var pnlReport = advancedTab.add('panel', undefined, 'Report Preferences');
    pnlReport.orientation = 'column';
    pnlReport.alignChildren = 'left';
    pnlReport.alignment = 'fill';
    pnlReport.margins = [15, 15, 15, 15];
    var cbReportImages = pnlReport.add('checkbox', undefined, 'Include image placement results'); cbReportImages.value = true;
    var cbReportOverset = pnlReport.add('checkbox', undefined, 'Include overset text summary'); cbReportOverset.value = true;
    var cbReportCounts = pnlReport.add('checkbox', undefined, 'Include session/topic counts'); cbReportCounts.value = true;

    // (Profiles feature removed)

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
        var allSettings = collectAllSettings(contentTab, contentTab, formattingTab, formattingTab, advancedTab);
        if (exportAllSettings(allSettings)) { alert('All settings exported successfully!'); }
    };

    // (Profiles UI and handlers removed)

    // (Defaults section removed)

    // Expose advanced controls for potential future use
    advancedTab.reportIncludeImages = cbReportImages;
    advancedTab.reportIncludeOverset = cbReportOverset;
    advancedTab.reportIncludeCounts = cbReportCounts;

    // (No profiles to initialize)
    
    // Dialog buttons with credit
    var grpBtns = dlg.add('group');
    grpBtns.orientation = 'row';
    grpBtns.alignment = 'fill';

    // Add credit on the left side
    var txtCredit = grpBtns.add('statictext', undefined, 'noorr @2025 | v2.0');
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
    
    // Pre-measure all tabs and set dialog to the largest required size to avoid overflow on other tabs
    function sizeDialogToLargestTab(dlgWin, tabs) {
        try {
            var originalSel = tabs.selection;
            var maxW = 0, maxH = 0;

            function relayout(win) {
                try {
                    if (win && win.layout) {
                        win.layout.layout(true);
                        win.layout.resize();
                        // extra pass to stabilize
                        win.layout.layout(true);
                    }
                    if (typeof win.update === 'function') win.update();
                } catch (e) {}
            }

            for (var i = 0; i < tabs.children.length; i++) {
                tabs.selection = tabs.children[i];
                // Let the tab apply its conditional UI
                try { if (typeof tabs.children[i].syncUI === 'function') tabs.children[i].syncUI(); } catch (eSync) {}
                relayout(dlgWin);

                var b = dlgWin.bounds; // [l,t,r,b]
                var w = b.right - b.left;
                var h = b.bottom - b.top;
                if (w > maxW) maxW = w;
                if (h > maxH) maxH = h;
            }

            // Restore original selection
            tabs.selection = originalSel || tabs.children[0];

            if (maxW > 0 && maxH > 0) {
                dlgWin.preferredSize = [maxW, maxH];
                try { dlgWin.minimumSize = [maxW, maxH]; } catch (eMin) {}
                relayout(dlgWin);
            }
        } catch (eOuter) {
            // Fallback to a single relayout
            try { if (dlgWin && dlgWin.layout) { dlgWin.layout.layout(true); dlgWin.layout.resize(); } } catch (ignored) {}
        }
    }

    // Compute and apply maximum dialog size across all tabs before showing
    sizeDialogToLargestTab(dlg, tabGroup);

    if (dlg.show() !== 1) return null;
    
    // Collect all settings and return using merged tabs
    return collectAllSettings(contentTab, contentTab, formattingTab, formattingTab, advancedTab);
}

// CAD Info Tab: displays version info and CSV detection summary
function setupCADInfoTab(tab, csvInfo, analysis, onApplySuggested) {
    // Make sections full width for consistency
    tab.alignChildren = 'fill';
    tab.margins = [15, 15, 15, 15];

    // CSV File Analysis
    var pnlCSV = tab.add('panel', undefined, 'CSV File Analysis');
    pnlCSV.margins = [15, 15, 15, 15];
    pnlCSV.alignChildren = 'fill';
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

    // Detected fields list with dynamic height based on count
    var grpFields = pnlCSV.add('group');
    grpFields.orientation = 'column';
    grpFields.alignChildren = 'fill';
    grpFields.alignment = 'fill';
    grpFields.spacing = 6;
    var stDetected = grpFields.add('statictext', undefined, 'Detected fields:');
    try { stDetected.alignment = 'left'; } catch (e) {}
    var fieldsArr = (csvInfo && csvInfo.fieldsList) ? csvInfo.fieldsList : [];
    var lbFields = grpFields.add('listbox', undefined, fieldsArr);
    lbFields.alignment = 'fill';
    // Dynamic vertical sizing within sane bounds
    try {
        var count = fieldsArr && fieldsArr.length ? fieldsArr.length : 0;
        var minRows = 3, maxRows = 12;
        var rows = Math.max(minRows, Math.min(maxRows, count));
        var rowHeight = 18; // approx per item height
        lbFields.preferredSize.height = rows * rowHeight + 6;
    } catch (e) {
        // fallback
        lbFields.preferredSize.height = 180;
    }
    try { lbFields.preferredSize.width = 520; } catch (e) {}

    // Template Analysis
    var pnlTemplate = tab.add('panel', undefined, 'Template Analysis');
    pnlTemplate.margins = [15, 15, 15, 15];
    pnlTemplate.alignChildren = 'fill';
    function yn(vTrue) { return vTrue ? 'Yes' : 'No'; }
    var g1 = pnlTemplate.add('group'); g1.orientation = 'row'; g1.spacing = 10; g1.alignment = 'left'; var g1l = g1.add('statictext', undefined, 'Supports table layout:'); g1l.preferredSize.width = 200; var g1v = g1.add('statictext', undefined, yn(analysis && analysis.supportsTable)); g1v.alignment = 'left';
    var g2 = pnlTemplate.add('group'); g2.orientation = 'row'; g2.spacing = 10; g2.alignment = 'left'; var g2l = g2.add('statictext', undefined, 'Supports indie layout:'); g2l.preferredSize.width = 200; var g2v = g2.add('statictext', undefined, yn(analysis && analysis.supportsIndependent)); g2v.alignment = 'left';
    var g3 = pnlTemplate.add('group'); g3.orientation = 'row'; g3.spacing = 10; g3.alignment = 'left'; var g3l = g3.add('statictext', undefined, 'Image automation available:'); g3l.preferredSize.width = 200; var g3v = g3.add('statictext', undefined, yn(analysis && analysis.imageAutomationAvailable)); g3v.alignment = 'left';

    // CSV format help (moved from Content -> Chairpersons to Overview)
    var pnlCSVHelp = tab.add('panel', undefined, 'CSV Format Help');
    pnlCSVHelp.margins = [5, 5, 5, 5];  // Reduced margins from 15 to 5
    pnlCSVHelp.alignChildren = 'fill';
    var stCSVHelp = pnlCSVHelp.add('statictext', undefined, 'Use double pipe (||) to separate chairpersons in your CSV file.\nExample: "John Smith||Jane Doe||Robert Johnson"', {multiline: true});
    stCSVHelp.alignment = 'fill';
    stCSVHelp.preferredSize.height = 32;  // Fixed height for exactly 2 lines
    try { stCSVHelp.graphics.font = ScriptUI.newFont(stCSVHelp.graphics.font.name, ScriptUI.FontStyle.ITALIC, 10); } catch (e) {}
}

function setupChairpersonsTab(tab, defaultSettings) {
    tab.alignChildren = 'fill';
    tab.margins = [15, 15, 15, 15];
    
    // Layout mode selection
    var pnlDesc = tab.add('panel', undefined, 'Chairpersons');
    pnlDesc.margins = [15, 15, 15, 15];
    pnlDesc.alignChildren = 'fill';
    
    var grpMode = pnlDesc.add('group'); 
    grpMode.orientation = 'row';
    grpMode.spacing = 10;
    grpMode.add('statictext', undefined, 'Layout:').preferredSize.width = 80;
    // Radios inline with label
    var rbInline = grpMode.add('radiobutton', undefined, 'Inline');
    var rbGrid = grpMode.add('radiobutton', undefined, 'Grid');
    // Keep dropdown for backward compatibility (hidden)
    var ddMode = grpMode.add('dropdownlist', undefined, ['Inline (single frame)', 'Separate frames grid']);
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
    // Store references for collectors/updaters
    tab.ddMode = ddMode;
    tab.rbInline = rbInline;
    tab.rbGrid = rbGrid;
    ddMode.selection = defaultSettings.chairMode === 'grid' ? 1 : 0;
    ddMode.preferredSize.width = 200;
    
    // Inline options (moved inside Chairpersons panel)
    var pnlInline = pnlDesc.add('panel', undefined, 'Inline Settings');
    pnlInline.alignment = 'fill';
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
    
    // Grid options (moved inside Chairpersons panel)
    var pnlGrid = pnlDesc.add('panel', undefined, 'Grid Settings');
    pnlGrid.alignment = 'fill';
    pnlGrid.orientation = 'column'; 
    pnlGrid.alignChildren = 'left';
    pnlGrid.margins = [15, 15, 15, 15];
    
    var grpOrder = pnlGrid.add('group'); 
    grpOrder.orientation = 'row';
    grpOrder.spacing = 10;
    grpOrder.add('statictext', undefined, 'Arrangement:').preferredSize.width = 80;
    var ddOrder = grpOrder.add('dropdownlist', undefined, ['Row-first', 'Column-first']);
    ddOrder.selection = defaultSettings.chairOrder === 'col' ? 1 : 0;
    ddOrder.preferredSize.width = 150;
    
    var grpDims = pnlGrid.add('group'); 
    grpDims.orientation = 'row';
    grpDims.spacing = 10;
    // Dimensions: [cols] cols × [rows] rows
    grpDims.add('statictext', undefined, 'Dimensions:').preferredSize.width = 80;
    var etCols = grpDims.add('edittext', undefined, defaultSettings.chairColumns || '2'); 
    etCols.characters = 4;
    grpDims.add('statictext', undefined, 'cols ×');
    var etRows = grpDims.add('edittext', undefined, defaultSettings.chairRows || '2'); 
    etRows.characters = 4;
    grpDims.add('statictext', undefined, 'rows');
    var grpSpacing = pnlGrid.add('group');
    grpSpacing.orientation = 'row';
    grpSpacing.spacing = 10;
    grpSpacing.alignment = 'left';
    grpSpacing.add('statictext', undefined, 'Spacing:').preferredSize.width = 80;
    // Chairpersons spacing: numeric-only fields (no unit suffix) + units dropdown
    var etColSpace = grpSpacing.add('edittext', undefined, String(defaultSettings.chairColSpacing || 8)); etColSpace.characters = 4;
    grpSpacing.add('statictext', undefined, 'horiz');
    var etRowSpace = grpSpacing.add('edittext', undefined, String(defaultSettings.chairRowSpacing || 4)); etRowSpace.characters = 4;
    grpSpacing.add('statictext', undefined, 'vert');
    var ddChairUnits = grpSpacing.add('dropdownlist', undefined, ['pt', 'mm', 'cm', 'px']);
    ddChairUnits.selection = defaultSettings.chairUnits ? indexOfExact(['pt', 'mm', 'cm', 'px'], defaultSettings.chairUnits) : 0;

    // Numeric-only validation for spacing fields
    etColSpace.onChanging = function(){ 
        // Allow numbers and decimal points only
        var cleanText = this.text.replace(/[^0-9\.]/g, '');
        // Ensure only one decimal point
        var parts = cleanText.split('.');
        if (parts.length > 2) {
            cleanText = parts[0] + '.' + parts.slice(1).join('');
        }
        this.text = cleanText;
    };
    etRowSpace.onChanging = function(){ 
        // Allow numbers and decimal points only
        var cleanText = this.text.replace(/[^0-9\.]/g, '');
        // Ensure only one decimal point
        var parts = cleanText.split('.');
        if (parts.length > 2) {
            cleanText = parts[0] + '.' + parts.slice(1).join('');
        }
        this.text = cleanText;
    };

    
    var grpCenter = pnlGrid.add('group');
    grpCenter.orientation = 'row';
    grpCenter.spacing = 10;
    var cbCenter = grpCenter.add('checkbox', undefined, 'Center grid horizontally');
    cbCenter.value = defaultSettings.chairCenterGrid || false;
    
    // Image automation controls - moved back to grid configuration as they're related
    var grpImageEnable = pnlGrid.add('group');
    grpImageEnable.orientation = 'row';
    grpImageEnable.spacing = 10;
    var cbEnableImages = grpImageEnable.add('checkbox', undefined, 'Auto image placement');
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
    tab.ddUnits = ddChairUnits; // Store the chair units dropdown reference
    
    // Make syncUI accessible from outside
    // helpers to enforce panels fully collapse when hidden
    function ensureCollapsible(panel) {
        try { if (panel && panel.minimumSize) { panel.minimumSize.height = 0; } } catch (e) {}
    }
    function setCollapsed(panel, collapsed) {
        try {
            if (!panel) return;
            if (collapsed) {
                if (panel.minimumSize) panel.minimumSize.height = 0;
                if (panel.maximumSize) panel.maximumSize.height = 0;
                panel.visible = false;
            } else {
                if (panel.maximumSize) panel.maximumSize.height = 10000;
                panel.visible = true;
            }
        } catch (e) {}
    }
    // small helper to force relayout up to the top-most window
    function forceRelayout(ctrl) {
        try {
            var root = ctrl;
            // climb to top-most parent (Window)
            while (root && root.parent) root = root.parent;
            if (root && root.layout) {
                root.layout.layout(true);
                root.layout.resize();
                // extra layout pass to stabilize
                root.layout.layout(true);
            }
            if (typeof root.update === 'function') root.update();
        } catch (e) {}
    }

    tab.syncUI = function() {
        var gridMode = ddMode.selection.index === 1;
        // Contextual show/hide to reduce clutter
        if (pnlInline) { pnlInline.enabled = !gridMode; setCollapsed(pnlInline, gridMode); }
        if (pnlGrid) { pnlGrid.enabled = gridMode; setCollapsed(pnlGrid, !gridMode); }
        if (gridMode) {
            var rowFirst = ddOrder.selection.index === 0;
            if (etCols) etCols.enabled = rowFirst;
            if (etRows) etRows.enabled = !rowFirst;
        }
        // Sync image controls - only available in grid mode
        syncImageUI();
        // Force relayout so hidden panels don't leave empty space
        forceRelayout(tab);
    };

    // Ensure collapsible behavior for panels and set initial layout state
    ensureCollapsible(pnlInline);
    ensureCollapsible(pnlGrid);
    // Keep dropdown changes in sync too
    try { ddMode.onChange = function(){ if (typeof tab.syncUI === 'function') { tab.syncUI(); forceRelayout(tab); } }; } catch (e) {}
    try { if (typeof tab.syncUI === 'function') { tab.syncUI(); forceRelayout(tab); } } catch (e) {}

    // Lightweight input validation and sanitization
    try {
        if (etCols && typeof etCols.onChanging === 'undefined') {
            etCols.onChanging = function(){ this.text = this.text.replace(/[^0-9]/g, ''); };
        }
        if (etRows && typeof etRows.onChanging === 'undefined') {
            etRows.onChanging = function(){ this.text = this.text.replace(/[^0-9]/g, ''); };
        }
        if (etColSpace && typeof etColSpace.onChanging === 'undefined') {
            etColSpace.onChanging = function(){ 
                // Allow numbers and decimal points only
                var cleanText = this.text.replace(/[^0-9\.]/g, '');
                // Ensure only one decimal point
                var parts = cleanText.split('.');
                if (parts.length > 2) {
                    cleanText = parts[0] + '.' + parts.slice(1).join('');
                }
                this.text = cleanText;
            };
        }
        if (etRowSpace && typeof etRowSpace.onChanging === 'undefined') {
            etRowSpace.onChanging = function(){ 
                // Allow numbers and decimal points only
                var cleanText = this.text.replace(/[^0-9\.]/g, '');
                // Ensure only one decimal point
                var parts = cleanText.split('.');
                if (parts.length > 2) {
                    cleanText = parts[0] + '.' + parts.slice(1).join('');
                }
                this.text = cleanText;
            };
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
    var pnlDesc = tab.add('panel', undefined, 'Topics');
    pnlDesc.margins = [15, 15, 15, 15];
    pnlDesc.alignChildren = 'fill';
    
    // Layout selector (moved out of the 'Layout' panel to top of Topics)
    var grpLayout = pnlDesc.add('group');
    grpLayout.orientation = 'row';
    grpLayout.spacing = 12;
    grpLayout.alignment = 'left';
    var stLayout = grpLayout.add('statictext', undefined, 'Layout:');
    stLayout.preferredSize.width = 80;
    var rbTable = grpLayout.add('radiobutton', undefined, 'Table');
    var rbIndependent = grpLayout.add('radiobutton', undefined, 'Indie Row');
    // Enable radios only if layout is available; set hover help for disabled
    try {
        rbTable.enabled = !!hasTableLayout;
        rbTable.helpTip = hasTableLayout ? '' : 'Table layout disabled: template has no topicsTable label.';
    } catch (e) {}
    try {
        rbIndependent.enabled = !!hasIndependentLayout;
        rbIndependent.helpTip = hasIndependentLayout ? '' : 'Independent rows disabled: template missing topicTime/title/speaker frames.';
    } catch (e) {}

    // Table layout options (moved inside Topics panel)
    var pnlTable = pnlDesc.add('panel', undefined, 'Table Settings');
    pnlTable.alignChildren = 'left';
    pnlTable.alignment = 'fill';
    pnlTable.margins = [15, 15, 15, 15];
    
    var grpHeader = pnlTable.add('group');
    grpHeader.orientation = 'row';
    grpHeader.spacing = 10;
    var cbHeader = grpHeader.add('checkbox', undefined, 'Include Table Header Row');
    cbHeader.value = defaultSettings.topicIncludeHeader !== false;

    // Independent layout options (moved inside Topics panel)
    var pnlIndependent = pnlDesc.add('panel', undefined, 'Indie Row Settings');
    pnlIndependent.alignChildren = 'left';
    pnlIndependent.alignment = 'fill';
    pnlIndependent.margins = [15, 15, 15, 15];
    
    var grpSpace = pnlIndependent.add('group');
    grpSpace.orientation = 'row';
    grpSpace.spacing = 10;
    grpSpace.add('statictext', undefined, 'Vertical spacing:').preferredSize.width = 100;
    var etSpacing = grpSpace.add('edittext', undefined, String(defaultSettings.topicVerticalSpacing || 4)); 
    etSpacing.characters = 4;
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
    
    // strong relayout helper
    function forceRelayout(ctrl) {
        try { var root = ctrl; while (root && root.parent) root = root.parent; if (root && root.layout) { root.layout.layout(true); root.layout.resize(); } } catch (e) {}
    }
    function ensureCollapsible(panel) { try { if (panel && panel.minimumSize) panel.minimumSize.height = 0; } catch (e) {} }
    function setCollapsed(panel, collapsed) {
        try {
            if (!panel) return;
            if (collapsed) {
                if (panel.minimumSize) panel.minimumSize.height = 0;
                if (panel.maximumSize) panel.maximumSize.height = 0;
                panel.visible = false;
            } else {
                if (panel.maximumSize) panel.maximumSize.height = 10000;
                panel.visible = true;
            }
        } catch (e) {}
    }

    function syncUI() {
        var tableMode = rbTable.value;
        pnlTable.enabled = tableMode;
        pnlIndependent.enabled = !tableMode;
        // Contextual show/hide with real collapse
        setCollapsed(pnlTable, !tableMode);
        setCollapsed(pnlIndependent, tableMode);
        // Force relayout so hidden panels don't leave empty space
        forceRelayout(tab);
    }
    
    rbTable.onClick = function(){ syncUI(); forceRelayout(tab); };
    rbIndependent.onClick = function(){ syncUI(); forceRelayout(tab); };
    // Numeric-only validation for topic spacing (no key/wheel stepping)
    try {
        if (etSpacing) {
            etSpacing.onChanging = function(){ 
                // Allow numbers and decimal points only
                var cleanText = this.text.replace(/[^0-9\.]/g, '');
                // Ensure only one decimal point
                var parts = cleanText.split('.');
                if (parts.length > 2) {
                    cleanText = parts[0] + '.' + parts.slice(1).join('');
                }
                this.text = cleanText;
            };
        }
    } catch (e) {}
    // Removed inline disabled notes; reasons now shown as helpTips on disabled radios
    // Ensure collapsible panels and perform initial layout
    ensureCollapsible(pnlTable);
    ensureCollapsible(pnlIndependent);
    // Set initial radio according to defaults/template
    try {
        var initialMode = (defaultSettings && defaultSettings.topicMode) ? defaultSettings.topicMode : (hasTableLayout ? 'table' : 'independent');
        if (initialMode === 'table' && rbTable.enabled) {
            rbTable.value = true; rbIndependent.value = false;
        } else if (initialMode === 'independent' && rbIndependent.enabled) {
            rbIndependent.value = true; rbTable.value = false;
        } else if (rbTable.enabled) {
            rbTable.value = true; rbIndependent.value = false;
        } else {
            rbIndependent.value = true; rbTable.value = false;
        }
    } catch (e) {}
    syncUI();
    forceRelayout(tab);
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
    if (tab.etColSpace) tab.etColSpace.text = String(settings.colSpacing || 8);
    if (tab.etRowSpace) tab.etRowSpace.text = String(settings.rowSpacing || 4);
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
    if (tab.etSpacing) tab.etSpacing.text = String(settings.verticalSpacing || 4);
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

function collectAllSettings(chairTab, topicTab, lineBreakTab, formattingTab, advancedTab) {
    // Defensive helpers and defaults
    var defs = getDefaultSettings();
    function selIndex(dd, defIdx) { try { return (dd && dd.selection) ? dd.selection.index : defIdx; } catch (e) { return defIdx; } }
    function selText(dd, defTxt) { try { return (dd && dd.selection) ? dd.selection.text : defTxt; } catch (e) { return defTxt; } }
    function getText(et, defTxt) { try { return (et && typeof et.text !== 'undefined') ? et.text : defTxt; } catch (e) { return defTxt; } }
    function getBool(cb, defVal) { try { return (cb && typeof cb.value !== 'undefined') ? cb.value : defVal; } catch (e) { return defVal; } }
    function toIntSafe(txt, defVal) { var n = parseInt(txt, 10); if (isNaN(n)) n = defVal; return n; }

    chairTab = chairTab || {};
    topicTab = topicTab || {};
    lineBreakTab = lineBreakTab || {};

    // Collect chairperson settings (defensive)
    var modeIdx = selIndex(chairTab.ddMode, defs.chairMode === 'inline' ? 0 : 1);
    var orderIdx = selIndex(chairTab.ddOrder, defs.chairOrder === 'row' ? 0 : 1);
    var unitsTxt = selText(chairTab.ddUnits, defs.chairUnits);
    var colSpaceTxt = getText(chairTab.etColSpace, '' + (defs.chairColSpacing || 8));
    var rowSpaceTxt = getText(chairTab.etRowSpace, '' + (defs.chairRowSpacing || 4));
    
    // Parse numeric values from text fields
    var colSpaceNum = parseFloat(colSpaceTxt) || 8;
    var rowSpaceNum = parseFloat(rowSpaceTxt) || 4;
    
    var chairOptions = {
        mode: (modeIdx === 0) ? 'inline' : 'grid',
        inlineSeparator: getBool(chairTab.rbLines, defs.chairInlineSeparator === 'linebreak') ? 'linebreak' : 'comma',
        order: (orderIdx === 0) ? 'row' : 'col',
        columns: clampInt(toIntSafe(getText(chairTab.etCols, defs.chairColumns), defs.chairColumns), 1, 99),
        rows: clampInt(toIntSafe(getText(chairTab.etRows, defs.chairRows), defs.chairRows), 1, 99),
        colSpacing: colSpaceNum,
        rowSpacing: rowSpaceNum,
        colSpacingPt: convertToPoints(colSpaceNum, unitsTxt),
        rowSpacingPt: convertToPoints(rowSpaceNum, unitsTxt),
        units: unitsTxt,
        centerGrid: getBool(chairTab.cbCenter, !!defs.chairCenterGrid),
        enableImages: getBool(chairTab.cbEnableImages, false),
        imageFolder: getText(chairTab.etImageFolder, ''),
        imageFitting: selText(chairTab.ddImageFitting, 'Use Template Default')
    };

    // Collect topic settings (defensive)
    var topicModeTable = getBool(topicTab.rbTable, (defs.topicMode || 'table') === 'table');
    var topicUnits = selText(topicTab.ddSpaceUnits, defs.topicUnits);
    var spacingTxt = getText(topicTab.etSpacing, '' + (defs.topicVerticalSpacing || 4));
    var spacingNum = parseFloat(spacingTxt) || 4;
    
    var topicOptions = {
        mode: topicModeTable ? 'table' : 'independent',
        includeHeader: getBool(topicTab.cbHeader, !!defs.topicIncludeHeader),
        verticalSpacing: spacingNum,
        verticalSpacingPt: convertToPoints(spacingNum, topicUnits),
        units: topicUnits
    };

    // Collect line break settings (defensive)
    var lineBreaks = {};
    if (lineBreakTab.sessionControls) {
        for (var field in lineBreakTab.sessionControls) {
            var sc = lineBreakTab.sessionControls[field];
            lineBreaks[field] = { enabled: getBool(sc && sc.checkbox, false), character: getText(sc && sc.charInput, '|') };
        }
    }
    if (lineBreakTab.chairpersonControls) {
        for (var field2 in lineBreakTab.chairpersonControls) {
            var cc = lineBreakTab.chairpersonControls[field2];
            lineBreaks[field2] = { enabled: getBool(cc && cc.checkbox, false), character: getText(cc && cc.charInput, '|') };
        }
    }
    if (lineBreakTab.topicControls) {
        for (var field3 in lineBreakTab.topicControls) {
            var tc = lineBreakTab.topicControls[field3];
            lineBreaks[field3] = { enabled: getBool(tc && tc.checkbox, false), character: getText(tc && tc.charInput, '|') };
        }
    }
    var lineBreakOptions = { lineBreaks: lineBreaks };

    // Collect styles (Phase 1 - paragraph styles)
    function val(dd) { 
        return (dd && dd.selection) ? dd.selection.text : '';
    }
    var stylesOptions = {
        table: {
            tableStyle: val(formattingTab && formattingTab.ddTableStyle),
            cellStyle: val(formattingTab && formattingTab.ddCellStyle)
        },
        session: {
            titlePara: val(formattingTab && formattingTab.ddSessionTitle),
            timePara: val(formattingTab && formattingTab.ddSessionTime),
            noPara: val(formattingTab && formattingTab.ddSessionNo)
        },
        chair: {
            style: val(formattingTab && formattingTab.ddChair)
        },
        topicsIndependent: {
            timePara: val(formattingTab && formattingTab.ddTopicTime),
            titlePara: val(formattingTab && formattingTab.ddTopicTitle),
            speakerPara: val(formattingTab && formattingTab.ddTopicSpeaker)
        }
    };

    // Collect report options from advanced tab
    var reportOptions = {
        includeImages: getBool(advancedTab && advancedTab.reportIncludeImages, true),
        includeOverset: getBool(advancedTab && advancedTab.reportIncludeOverset, true),
        includeCounts: getBool(advancedTab && advancedTab.reportIncludeCounts, true)
    };

    return {
        chairOptions: chairOptions,
        topicOptions: topicOptions,
        lineBreakOptions: lineBreakOptions,
        stylesOptions: stylesOptions,
        reportOptions: reportOptions
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

// Ensure document measurement units are set to points for accurate positioning
function ensurePointUnits() {
    var doc = app.activeDocument;
    var originalHUnits = doc.viewPreferences.horizontalMeasurementUnits;
    var originalVUnits = doc.viewPreferences.verticalMeasurementUnits;
    
    // Set to points
    doc.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.POINTS;
    doc.viewPreferences.verticalMeasurementUnits = MeasurementUnits.POINTS;
    
    return {
        horizontal: originalHUnits,
        vertical: originalVUnits
    };
}

// Restore original document measurement units
function restoreUnits(originalUnits) {
    try {
        var doc = app.activeDocument;
        doc.viewPreferences.horizontalMeasurementUnits = originalUnits.horizontal;
        doc.viewPreferences.verticalMeasurementUnits = originalUnits.vertical;
    } catch (e) {
        // Ignore restoration errors
    }
}

// Simplified and more reliable unit conversion function
function convertToPoints(numericValue, unit) {
    var n = parseFloat(numericValue);
    if (isNaN(n)) n = 0;
    
    var result;
    switch (unit) {
        case 'mm': result = n * 2.834645669;  // 1mm = 2.834645669 points
            break;
        case 'cm': result = n * 28.34645669;  // 1cm = 28.34645669 points
            break;
        case 'px': result = n;         // 1px = 1 point
        case 'pt':
        default: result = n;                  // points remain as points
            break;
    }
    
    return result;
}

// Helper function to safely apply paragraph style to a text frame
function applyParagraphStyleToFrame(textFrame, styleName) {
    if (!textFrame) return;
    
    // Handle "— None —" selection - break style linkage while preserving formatting
    if (!styleName || styleName === '' || styleName === '— None —') {
        
        if (textFrame.constructor.name !== "TextFrame") return;
        
        try {
            // First, detect what style is currently applied
            var currentStyleName = null;
            var hasUserStyle = false;
            
            // Check current style via parentStory method (multi-method fallback approach)
            try {
                if (textFrame.parentStory && textFrame.parentStory.paragraphs.length > 0) {
                    var currentStyle = textFrame.parentStory.paragraphs[0].appliedParagraphStyle;
                    if (currentStyle && currentStyle.name) {
                        currentStyleName = currentStyle.name;
                        
                        // Check if it's a user-created style (not system style)
                        if (currentStyleName !== '[Basic Paragraph]' && 
                            currentStyleName !== '[No Paragraph Style]' && 
                            currentStyleName.charAt(0) !== '[') {
                            hasUserStyle = true;
                        }
                    }
                }
            } catch (e) {
                // Fallback to texts method if parentStory fails
                try {
                    if (textFrame.texts && textFrame.texts.length > 0) {
                        var textStyle = textFrame.texts[0].appliedParagraphStyle;
                        if (textStyle && textStyle.name) {
                            currentStyleName = textStyle.name;
                            if (currentStyleName !== '[Basic Paragraph]' && 
                                currentStyleName !== '[No Paragraph Style]' && 
                                currentStyleName.charAt(0) !== '[') {
                                hasUserStyle = true;
                            }
                        }
                    }
                } catch (e2) {}
            }
            
            // Apply the correct behavior based on current state
            if (!hasUserStyle) {
                // Case 1: No user style preassigned - do nothing, preserve current characteristics
                return;
            } else {
                // Case 2: User style preassigned - break link but preserve ALL formatting
                // InDesign's [No Paragraph Style] clears overrides even via script, so we need
                // to capture formatting first, then apply [No Paragraph Style], then restore as overrides
                try {
                    if (textFrame.parentStory && textFrame.parentStory.paragraphs.length > 0) {
                        var paragraphs = textFrame.parentStory.paragraphs;
                        var doc = app.documents[0];
                        var noParaStyle = doc.paragraphStyles.item("[No Paragraph Style]");
                        
                        // Capture all formatting properties BEFORE breaking the link
                        var savedProps = [];
                        for (var i = 0; i < paragraphs.length; i++) {
                            try {
                                var para = paragraphs[i];
                                var props = {
                                    // Paragraph formatting
                                    pointSize: para.pointSize,
                                    leading: para.leading,
                                    leftIndent: para.leftIndent,
                                    rightIndent: para.rightIndent,
                                    firstLineIndent: para.firstLineIndent,
                                    spaceAfter: para.spaceAfter,
                                    spaceBefore: para.spaceBefore,
                                    justification: para.justification,
                                    fillColor: para.fillColor,
                                    strokeColor: para.strokeColor,
                                    // Character formatting (from first character to preserve)
                                    appliedFont: para.characters.length > 0 ? para.characters[0].appliedFont : null,
                                    fontStyle: para.characters.length > 0 ? para.characters[0].fontStyle : null
                                };
                                savedProps.push(props);
                            } catch (eProp) {
                                savedProps.push(null);
                            }
                        }
                        
                        // Now break the style link by applying [No Paragraph Style]
                        for (var j = 0; j < paragraphs.length; j++) {
                            try {
                                paragraphs[j].appliedParagraphStyle = noParaStyle;
                            } catch (eBreak) {}
                        }
                        
                        // Restore all captured formatting as local overrides
                        for (var k = 0; k < paragraphs.length && k < savedProps.length; k++) {
                            if (savedProps[k]) {
                                try {
                                    var para = paragraphs[k];
                                    var props = savedProps[k];
                                    
                                    // Restore paragraph properties
                                    if (props.pointSize !== undefined) para.pointSize = props.pointSize;
                                    if (props.leading !== undefined) para.leading = props.leading;
                                    if (props.leftIndent !== undefined) para.leftIndent = props.leftIndent;
                                    if (props.rightIndent !== undefined) para.rightIndent = props.rightIndent;
                                    if (props.firstLineIndent !== undefined) para.firstLineIndent = props.firstLineIndent;
                                    if (props.spaceAfter !== undefined) para.spaceAfter = props.spaceAfter;
                                    if (props.spaceBefore !== undefined) para.spaceBefore = props.spaceBefore;
                                    if (props.justification !== undefined) para.justification = props.justification;
                                    if (props.fillColor) para.fillColor = props.fillColor;
                                    if (props.strokeColor) para.strokeColor = props.strokeColor;
                                    
                                    // Restore character properties to all characters in paragraph
                                    if (para.characters.length > 0) {
                                        if (props.appliedFont) para.characters.everyItem().appliedFont = props.appliedFont;
                                        if (props.fontStyle) para.characters.everyItem().fontStyle = props.fontStyle;
                                    }
                                } catch (eRestore) {}
                            }
                        }
                    }
                } catch (e) {
                    // Fallback: try via texts collection if parentStory fails
                    try {
                        if (textFrame.texts && textFrame.texts.length > 0) {
                            var textObj = textFrame.texts[0];
                            var doc = app.documents[0];
                            var noParaStyle = doc.paragraphStyles.item("[No Paragraph Style]");
                            
                            if (textObj.paragraphs && textObj.paragraphs.length > 0) {
                                // Simple fallback - just break the link
                                textObj.paragraphs.everyItem().appliedParagraphStyle = noParaStyle;
                            }
                        }
                    } catch (eFallback) {}
                }
            }
        } catch (e) {}
        return;
    }
    
    // Regular style application (when not "— None —")
    if (textFrame.constructor.name !== "TextFrame") return;
    
    try {
        var doc = app.activeDocument;
        var paragraphStyle = doc.paragraphStyles.itemByName(styleName);
        
        if (paragraphStyle && paragraphStyle.isValid) {
            // Use parentStory method (works based on previous debug output)
            try {
                if (textFrame.parentStory && textFrame.parentStory.paragraphs.length > 0) {
                    textFrame.parentStory.paragraphs.everyItem().appliedParagraphStyle = paragraphStyle;
                }
            } catch (e1) {
                // Fallback to texts method
                try {
                    if (textFrame.texts.length > 0) {
                        textFrame.texts.everyItem().appliedParagraphStyle = paragraphStyle;
                    }
                } catch (e2) {}
            }
        }
    } catch (e) {
        $.writeln("Warning: Could not apply paragraph style '" + styleName + "' to text frame: " + e);
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

    // Fill session-level placeholders with line break support and paragraph styles
    if (placeholders.sessionTitle && session.title) {
        placeholders.sessionTitle.contents = session.title;
        // Apply paragraph style if specified
        if (stylesOptions && stylesOptions.session && stylesOptions.session.titlePara) {
            applyParagraphStyleToFrame(placeholders.sessionTitle, stylesOptions.session.titlePara);
        }
        if (lineBreakOptions.lineBreaks.SessionTitle && lineBreakOptions.lineBreaks.SessionTitle.enabled) {
            applyLineBreaksToFrame(placeholders.sessionTitle, lineBreakOptions.lineBreaks.SessionTitle.character);
        }
    }
    if (placeholders.sessionTime && session.time) {
        placeholders.sessionTime.contents = session.time;
        // Apply paragraph style if specified
        if (stylesOptions && stylesOptions.session && stylesOptions.session.timePara) {
            applyParagraphStyleToFrame(placeholders.sessionTime, stylesOptions.session.timePara);
        }
        if (lineBreakOptions.lineBreaks.SessionTime && lineBreakOptions.lineBreaks.SessionTime.enabled) {
            applyLineBreaksToFrame(placeholders.sessionTime, lineBreakOptions.lineBreaks.SessionTime.character);
        }
    }
    if (placeholders.sessionNo && session.no) {
        placeholders.sessionNo.contents = session.no;
        // Apply paragraph style if specified
        if (stylesOptions && stylesOptions.session && stylesOptions.session.noPara) {
            applyParagraphStyleToFrame(placeholders.sessionNo, stylesOptions.session.noPara);
        }
        if (lineBreakOptions.lineBreaks.SessionNo && lineBreakOptions.lineBreaks.SessionNo.enabled) {
            applyLineBreaksToFrame(placeholders.sessionNo, lineBreakOptions.lineBreaks.SessionNo.character);
        }
    }
    
    // Fill chairpersons
    if (placeholders.chairpersons && session.chairs) {
        populateChairs(page, placeholders.chairpersons, session.chairs, chairOptions, lineBreakOptions, stylesOptions);
    }

    // Fill topics
    if (topicOptions.mode === 'table' && placeholders.topicsTable) {
        populateTopicsAsTable(placeholders.topicsTable, session.topics, topicOptions, lineBreakOptions, stylesOptions);
    } else if (topicOptions.mode === 'independent' && hasIndependentTopicLayout) {
        populateTopicsAsIndependentRows(page, session.topics, topicOptions, lineBreakOptions, stylesOptions);
    }
}

function populateChairs(page, chairFrame, chairsRaw, opts, lineBreakOptions, stylesOptions) {
    // Ensure document is using points for accurate positioning
    var originalUnits = ensurePointUnits();
    
    try {
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
        
        // Apply paragraph style if specified
        if (stylesOptions && stylesOptions.chair && stylesOptions.chair.style) {
            applyParagraphStyleToFrame(chairFrame, stylesOptions.chair.style);
        }
        
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
            // Apply paragraph style if specified
            if (stylesOptions && stylesOptions.chair && stylesOptions.chair.style) {
                applyParagraphStyleToFrame(tf, stylesOptions.chair.style);
            }
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
    
    } finally {
        // Restore original document units
        restoreUnits(originalUnits);
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

function populateTopicsAsIndependentRows(page, topics, opts, lineBreakOptions, stylesOptions) {
    // Ensure document is using points for accurate positioning
    var originalUnits = ensurePointUnits();
    
    try {
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
                // Apply paragraph style if specified
                if (stylesOptions && stylesOptions.topicsIndependent && stylesOptions.topicsIndependent.timePara) {
                    applyParagraphStyleToFrame(tTime, stylesOptions.topicsIndependent.timePara);
                }
                if (lineBreakOptions.lineBreaks.topicTime && lineBreakOptions.lineBreaks.topicTime.enabled) {
                    applyLineBreaksToFrame(tTime, lineBreakOptions.lineBreaks.topicTime.character);
                }
            }
            if (tTopic) {
                tTopic.contents = topics[i].title;
                // Apply paragraph style if specified
                if (stylesOptions && stylesOptions.topicsIndependent && stylesOptions.topicsIndependent.titlePara) {
                    applyParagraphStyleToFrame(tTopic, stylesOptions.topicsIndependent.titlePara);
                }
                if (lineBreakOptions.lineBreaks.topicTitle && lineBreakOptions.lineBreaks.topicTitle.enabled) {
                    applyLineBreaksToFrame(tTopic, lineBreakOptions.lineBreaks.topicTitle.character);
                }
            }
            if (tSpeaker) {
                tSpeaker.contents = topics[i].speaker;
                // Apply paragraph style if specified
                if (stylesOptions && stylesOptions.topicsIndependent && stylesOptions.topicsIndependent.speakerPara) {
                    applyParagraphStyleToFrame(tSpeaker, stylesOptions.topicsIndependent.speakerPara);
                }
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
            // Apply paragraph style if specified
            if (stylesOptions && stylesOptions.topicsIndependent && stylesOptions.topicsIndependent.timePara) {
                applyParagraphStyleToFrame(timeClone, stylesOptions.topicsIndependent.timePara);
            }
            if (lineBreakOptions.lineBreaks.topicTime && lineBreakOptions.lineBreaks.topicTime.enabled) {
                applyLineBreaksToFrame(timeClone, lineBreakOptions.lineBreaks.topicTime.character);
            }
            timeClone.move(undefined, [0, targetY - timeClone.geometricBounds[0]]);
            
            // Clone and position Topic frame
            var topicClone = protoTopic.duplicate(page);
            topicClone.label = "topicClone_Topic";
            topicClone.visible = true;
            topicClone.contents = topics[i].title;
            // Apply paragraph style if specified
            if (stylesOptions && stylesOptions.topicsIndependent && stylesOptions.topicsIndependent.titlePara) {
                applyParagraphStyleToFrame(topicClone, stylesOptions.topicsIndependent.titlePara);
            }
            if (lineBreakOptions.lineBreaks.topicTitle && lineBreakOptions.lineBreaks.topicTitle.enabled) {
                applyLineBreaksToFrame(topicClone, lineBreakOptions.lineBreaks.topicTitle.character);
            }
            topicClone.move(undefined, [0, targetY - topicClone.geometricBounds[0]]);
            
            // Clone and position Speaker frame
            var speakerClone = protoSpeaker.duplicate(page);
            speakerClone.label = "topicClone_Speaker";
            speakerClone.visible = true;
            speakerClone.contents = topics[i].speaker;
            // Apply paragraph style if specified
            if (stylesOptions && stylesOptions.topicsIndependent && stylesOptions.topicsIndependent.speakerPara) {
                applyParagraphStyleToFrame(speakerClone, stylesOptions.topicsIndependent.speakerPara);
            }
            if (lineBreakOptions.lineBreaks.topicSpeaker && lineBreakOptions.lineBreaks.topicSpeaker.enabled) {
                applyLineBreaksToFrame(speakerClone, lineBreakOptions.lineBreaks.topicSpeaker.character);
            }
            speakerClone.move(undefined, [0, targetY - speakerClone.geometricBounds[0]]);
        }
    }
    
    } finally {
        // Restore original document units
        restoreUnits(originalUnits);
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
                style: ''
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

function exportReport(sessionsData, reportOptions) {
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
    
    // Set default values if reportOptions is not provided
    var includeOverset = reportOptions ? reportOptions.includeOverset : true;
    var includeImages = reportOptions ? reportOptions.includeImages : true;
    var includeCounts = reportOptions ? reportOptions.includeCounts : true;
    
    // Write the report header
    reportFile.writeln("AGENDA IMPORT REPORT");
    reportFile.writeln("====================");
    reportFile.writeln("Date: " + new Date().toLocaleString());
    reportFile.writeln("Sessions imported: " + sessionsData.length);
    reportFile.writeln("");
    
    // Check for overset text errors (only if enabled in preferences)
    if (includeOverset) {
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
        } catch (e) {
            // Test frame creation failed - continue without test frame
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
    
    } // End of overset text check conditional
    
    // Add image placement report if automation was used and if enabled in preferences
    if (includeImages && imageResults.enabled) {
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
    
    // Sessions report (only if enabled in preferences)
    if (includeCounts) {
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
    
    } // End of sessions report conditional
    
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
        selectDropdownByText(ddChair, s.chair.style);
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
                var sel = (dd.selection && dd.selection.text) ? dd.selection.text : null;
                dd.removeAll();
                for (var i = 0; i < itemsArr.length; i++) dd.add('item', itemsArr[i]);
                dd.selection = sel ? (function(){ for (var j = 0; j < dd.items.length; j++){ if(dd.items[j].text===sel) return j;} return 0;})() : 0;
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
            selectDropdownByText(tab.ddChair, stylesOptions.chair.style);
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

/* ---------------- DOCUMENT SETUP COPYING ---------------- */

/**
 * Safely get a document preference property with fallback for older InDesign versions
 */
function safeGetDocProperty(doc, propertyName, fallbackValue) {
    try {
        if (doc.documentPreferences.hasOwnProperty(propertyName)) {
            return doc.documentPreferences[propertyName];
        }
        return fallbackValue;
    } catch (e) {
        $.writeln("Warning: Property '" + propertyName + "' not supported in this InDesign version: " + e);
        return fallbackValue;
    }
}

/**
 * Safely set a document preference property with fallback for older InDesign versions
 */
function safeSetDocProperty(doc, propertyName, value) {
    try {
        if (doc.documentPreferences.hasOwnProperty(propertyName)) {
            doc.documentPreferences[propertyName] = value;
            return true;
        }
        $.writeln("Warning: Property '" + propertyName + "' not supported in this InDesign version");
        return false;
    } catch (e) {
        $.writeln("Warning: Could not set property '" + propertyName + "': " + e);
        return false;
    }
}

/**
 * Copy comprehensive document setup options from source template to new document
 * 
 * Includes:
 * - Page size, orientation, margins, bleeds, slugs
 * - Facing pages, binding, and page direction settings
 * - View preferences (measurement units, rulers, guides visibility)
 * - Grid and baseline grid settings
 * - Guide preferences and behavior
 * - Text preferences and typography settings
 * - Print setup preferences
 * - Color management settings
 * - Transparency and pasteboard preferences
 * - Story preferences (optical margin, smart text reflow)
 * 
 * This ensures the new document inherits all formatting and display
 * preferences from the original template for consistent workflow.
 */
function copyDocumentSettings(sourceDoc, targetDoc) {
    try {
        // CRITICAL: Copy facing pages FIRST - this affects all other layout settings
        var sourceFacingPages = sourceDoc.documentPreferences.facingPages;
        var sourcePageBinding = sourceDoc.documentPreferences.pageBinding;
        var sourcePageDirection = safeGetDocProperty(sourceDoc, 'pageDirection', 'N/A');
        
        $.writeln("Source document setup:");
        $.writeln("  - Facing Pages: " + sourceFacingPages);
        $.writeln("  - Page Binding: " + sourcePageBinding);
        $.writeln("  - Page Direction: " + sourcePageDirection);
        
        // Set facing pages first as it affects margin calculations
        targetDoc.documentPreferences.facingPages = sourceFacingPages;
        
        // Copy page size and orientation
        targetDoc.documentPreferences.pageWidth = sourceDoc.documentPreferences.pageWidth;
        targetDoc.documentPreferences.pageHeight = sourceDoc.documentPreferences.pageHeight;
        targetDoc.documentPreferences.pageOrientation = sourceDoc.documentPreferences.pageOrientation;
        
        // Copy binding settings AFTER facing pages is set
        targetDoc.documentPreferences.pageBinding = sourcePageBinding;
        
        // Safely copy pageDirection (not available in all InDesign versions)
        if (sourcePageDirection !== 'N/A') {
            safeSetDocProperty(targetDoc, 'pageDirection', sourcePageDirection);
        }
        
        // Copy margins AFTER facing pages and binding are set (affects inside/outside margins)
        targetDoc.documentPreferences.documentMarginTop = sourceDoc.documentPreferences.documentMarginTop;
        targetDoc.documentPreferences.documentMarginBottom = sourceDoc.documentPreferences.documentMarginBottom;
        targetDoc.documentPreferences.documentMarginLeft = sourceDoc.documentPreferences.documentMarginLeft;
        targetDoc.documentPreferences.documentMarginRight = sourceDoc.documentPreferences.documentMarginRight;
        
        // Copy inside/outside margins if facing pages are enabled
        try {
            if (sourceFacingPages && sourceDoc.documentPreferences.hasOwnProperty('documentMarginInside')) {
                targetDoc.documentPreferences.documentMarginInside = sourceDoc.documentPreferences.documentMarginInside;
                targetDoc.documentPreferences.documentMarginOutside = sourceDoc.documentPreferences.documentMarginOutside;
                $.writeln("  ✓ Copied inside/outside margins for facing pages");
            }
        } catch (insideMarginErr) {
            $.writeln("Warning: Could not copy inside/outside margins: " + insideMarginErr);
        }
        
        // Copy additional document setup properties
        try {
            targetDoc.documentPreferences.pagesPerDocument = sourceDoc.documentPreferences.pagesPerDocument;
        } catch (pagesErr) {
            $.writeln("Warning: Could not copy pagesPerDocument: " + pagesErr);
        }
        
        try {
            targetDoc.documentPreferences.startPageNumber = sourceDoc.documentPreferences.startPageNumber;
        } catch (startErr) {
            $.writeln("Warning: Could not copy startPageNumber: " + startErr);
        }
        
        try {
            targetDoc.documentPreferences.columnCount = sourceDoc.documentPreferences.columnCount;
            targetDoc.documentPreferences.columnGutter = sourceDoc.documentPreferences.columnGutter;
        } catch (colErr) {
            $.writeln("Warning: Could not copy column settings: " + colErr);
        }
        
        // Copy additional layout and setup properties
        try {
            // Master page shuffling and item behavior
            targetDoc.documentPreferences.allowPageShuffle = sourceDoc.documentPreferences.allowPageShuffle;
            targetDoc.documentPreferences.allowMasterPageItemsToResize = sourceDoc.documentPreferences.allowMasterPageItemsToResize;
            targetDoc.documentPreferences.allowMasterPageItemsToMove = sourceDoc.documentPreferences.allowMasterPageItemsToMove;
        } catch (shuffleErr) {
            $.writeln("Warning: Could not copy master page behavior settings: " + shuffleErr);
        }
        
        // Copy document intent (print/web) - affects many layout behaviors
        try {
            targetDoc.documentPreferences.intent = sourceDoc.documentPreferences.intent;
            targetDoc.documentPreferences.outputIntent = sourceDoc.documentPreferences.outputIntent;
        } catch (intentErr) {
            $.writeln("Warning: Could not copy document intent: " + intentErr);
        }
        
        // Copy spread and page positioning settings
        try {
            // These properties affect how pages are displayed and arranged
            if (sourceDoc.documentPreferences.hasOwnProperty('spreadDirection')) {
                targetDoc.documentPreferences.spreadDirection = sourceDoc.documentPreferences.spreadDirection;
            }
            if (sourceDoc.documentPreferences.hasOwnProperty('pagePositioning')) {
                targetDoc.documentPreferences.pagePositioning = sourceDoc.documentPreferences.pagePositioning;
            }
        } catch (posErr) {
            $.writeln("Warning: Could not copy spread positioning settings: " + posErr);
        }
        
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
        try {
            targetDoc.documentPreferences.pageNumberingStartPage = sourceDoc.documentPreferences.pageNumberingStartPage;
            targetDoc.documentPreferences.pageNumberingSectionPrefix = sourceDoc.documentPreferences.pageNumberingSectionPrefix;
        } catch (numberingErr) {
            $.writeln("Warning: Could not copy page numbering settings: " + numberingErr);
        }
        
        // Copy view preferences (measurement units, rulers, guides visibility)
        try {
            targetDoc.viewPreferences.horizontalMeasurementUnits = sourceDoc.viewPreferences.horizontalMeasurementUnits;
            targetDoc.viewPreferences.verticalMeasurementUnits = sourceDoc.viewPreferences.verticalMeasurementUnits;
            targetDoc.viewPreferences.typographicMeasurementUnits = sourceDoc.viewPreferences.typographicMeasurementUnits;
            targetDoc.viewPreferences.rulerOrigin = sourceDoc.viewPreferences.rulerOrigin;
            targetDoc.viewPreferences.showRulers = sourceDoc.viewPreferences.showRulers;
            targetDoc.viewPreferences.showFrameEdges = sourceDoc.viewPreferences.showFrameEdges;
            targetDoc.viewPreferences.showMargins = sourceDoc.viewPreferences.showMargins;
            targetDoc.viewPreferences.showBaselines = sourceDoc.viewPreferences.showBaselines;
            targetDoc.viewPreferences.showDocumentGrid = sourceDoc.viewPreferences.showDocumentGrid;
            targetDoc.viewPreferences.showBaselines = sourceDoc.viewPreferences.showBaselines;
        } catch (viewErr) {
            $.writeln("Warning: Could not copy all view preferences: " + viewErr);
        }
        
        // Copy grid preferences
        try {
            targetDoc.gridPreferences.horizontalGridlineDivision = sourceDoc.gridPreferences.horizontalGridlineDivision;
            targetDoc.gridPreferences.horizontalGridlineSubdivision = sourceDoc.gridPreferences.horizontalGridlineSubdivision;
            targetDoc.gridPreferences.verticalGridlineDivision = sourceDoc.gridPreferences.verticalGridlineDivision;
            targetDoc.gridPreferences.verticalGridlineSubdivision = sourceDoc.gridPreferences.verticalGridlineSubdivision;
            targetDoc.gridPreferences.baselineStart = sourceDoc.gridPreferences.baselineStart;
            targetDoc.gridPreferences.baselineDivision = sourceDoc.gridPreferences.baselineDivision;
            targetDoc.gridPreferences.baselineShown = sourceDoc.gridPreferences.baselineShown;
            targetDoc.gridPreferences.documentGridShown = sourceDoc.gridPreferences.documentGridShown;
            targetDoc.gridPreferences.gridShown = sourceDoc.gridPreferences.gridShown;
        } catch (gridErr) {
            $.writeln("Warning: Could not copy grid preferences: " + gridErr);
        }
        
        // Copy guide preferences
        try {
            targetDoc.guidePreferences.guidesShown = sourceDoc.guidePreferences.guidesShown;
            targetDoc.guidePreferences.guidesLocked = sourceDoc.guidePreferences.guidesLocked;
            targetDoc.guidePreferences.guidesInBack = sourceDoc.guidePreferences.guidesInBack;
            targetDoc.guidePreferences.guidesSnapto = sourceDoc.guidePreferences.guidesSnapto;
            targetDoc.guidePreferences.rulerGuidesViewThreshold = sourceDoc.guidePreferences.rulerGuidesViewThreshold;
        } catch (guideErr) {
            $.writeln("Warning: Could not copy guide preferences: " + guideErr);
        }
        
        // Copy text preferences
        try {
            targetDoc.textPreferences.showInvisibles = sourceDoc.textPreferences.showInvisibles;
            targetDoc.textPreferences.justifyTextWraps = sourceDoc.textPreferences.justifyTextWraps;
            targetDoc.textPreferences.abutTextToTextWrap = sourceDoc.textPreferences.abutTextToTextWrap;
            targetDoc.textPreferences.linkTextFilesWhenImporting = sourceDoc.textPreferences.linkTextFilesWhenImporting;
            targetDoc.textPreferences.moveSystemClipboard = sourceDoc.textPreferences.moveSystemClipboard;
            targetDoc.textPreferences.useOpticalMarginAlignment = sourceDoc.textPreferences.useOpticalMarginAlignment;
            targetDoc.textPreferences.useParagraphLeading = sourceDoc.textPreferences.useParagraphLeading;
        } catch (textErr) {
            $.writeln("Warning: Could not copy text preferences: " + textErr);
        }
        
        // Copy print preferences (basic settings)
        try {
            targetDoc.printPreferences.pagePosition = sourceDoc.printPreferences.pagePosition;
            targetDoc.printPreferences.thumbnails = sourceDoc.printPreferences.thumbnails;
            targetDoc.printPreferences.useDocumentBleedToPrint = sourceDoc.printPreferences.useDocumentBleedToPrint;
            targetDoc.printPreferences.printBlankPages = sourceDoc.printPreferences.printBlankPages;
            targetDoc.printPreferences.printMasterPages = sourceDoc.printPreferences.printMasterPages;
            targetDoc.printPreferences.printNonprintingObjects = sourceDoc.printPreferences.printNonprintingObjects;
            targetDoc.printPreferences.printVisibleGuidesAndBaselines = sourceDoc.printPreferences.printVisibleGuidesAndBaselines;
        } catch (printErr) {
            $.writeln("Warning: Could not copy print preferences: " + printErr);
        }
        
        // Copy transparency preferences
        try {
            targetDoc.transparencyPreferences.blendingColorSpace = sourceDoc.transparencyPreferences.blendingColorSpace;
        } catch (transErr) {
            $.writeln("Warning: Could not copy transparency preferences: " + transErr);
        }
        
        // Copy story preferences
        try {
            targetDoc.storyPreferences.opticalMarginAlignment = sourceDoc.storyPreferences.opticalMarginAlignment;
            targetDoc.storyPreferences.opticalMarginSize = sourceDoc.storyPreferences.opticalMarginSize;
            targetDoc.storyPreferences.smartTextReflow = sourceDoc.storyPreferences.smartTextReflow;
        } catch (storyErr) {
            $.writeln("Warning: Could not copy story preferences: " + storyErr);
        }
        
        // Copy color management settings
        try {
            targetDoc.colorPreferences.cmykPolicy = sourceDoc.colorPreferences.cmykPolicy;
            targetDoc.colorPreferences.rgbPolicy = sourceDoc.colorPreferences.rgbPolicy;
            targetDoc.colorPreferences.enableColorManagement = sourceDoc.colorPreferences.enableColorManagement;
        } catch (colorErr) {
            $.writeln("Warning: Could not copy color management settings: " + colorErr);
        }
        
        // Copy pasteboard preferences
        try {
            targetDoc.pasteboardPreferences.pasteboardMargins = sourceDoc.pasteboardPreferences.pasteboardMargins;
            targetDoc.pasteboardPreferences.minimumSpaceAboveAndBelow = sourceDoc.pasteboardPreferences.minimumSpaceAboveAndBelow;
        } catch (pasteErr) {
            $.writeln("Warning: Could not copy pasteboard preferences: " + pasteErr);
        }
        
        $.writeln("\n=== DOCUMENT SETUP COPY VERIFICATION ===");
        $.writeln("Document Setup Properties Copied:");
        $.writeln("  ✓ Facing Pages: " + targetDoc.documentPreferences.facingPages + " (was: " + sourceFacingPages + ")");
        $.writeln("  ✓ Page Binding: " + targetDoc.documentPreferences.pageBinding + " (was: " + sourcePageBinding + ")");
        
        var currentPageDirection = safeGetDocProperty(targetDoc, 'pageDirection', 'N/A');
        if (sourcePageDirection !== 'N/A') {
            $.writeln("  ✓ Page Direction: " + currentPageDirection + " (was: " + sourcePageDirection + ")");
        } else {
            $.writeln("  • Page Direction: Not supported in this InDesign version");
        }
        
        $.writeln("  ✓ Page Size: " + targetDoc.documentPreferences.pageWidth + " x " + targetDoc.documentPreferences.pageHeight);
        $.writeln("  ✓ Page Orientation: " + targetDoc.documentPreferences.pageOrientation);
        
        try {
            $.writeln("  ✓ Pages Per Document: " + targetDoc.documentPreferences.pagesPerDocument);
            $.writeln("  ✓ Start Page Number: " + targetDoc.documentPreferences.startPageNumber);
            $.writeln("  ✓ Columns: " + targetDoc.documentPreferences.columnCount + " (gutter: " + targetDoc.documentPreferences.columnGutter + ")");
        } catch (debugErr) {
            $.writeln("  Note: Some additional properties could not be displayed");
        }
        
        $.writeln("  ✓ Measurement Units: " + targetDoc.viewPreferences.horizontalMeasurementUnits + " x " + targetDoc.viewPreferences.verticalMeasurementUnits);
        $.writeln("=== END DOCUMENT SETUP VERIFICATION ===");
        
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