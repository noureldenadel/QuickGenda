v2.0
//   Copy Parent page elements to new document
//   Phase 1 – Core CSV Import & Table Integration
    ✅ Modify script to detect if a labeled text frame already contains a table.
    If yes → inject CSV into the existing table (preserve styles).
    If no → create a new table from CSV.
    ✅ Implement logic to expand or shrink table to fit CSV (add/remove rows & columns).
    ✅ Fill table cell-by-cell with CSV data, without overwriting styling.
// add styling tab : workflow / panel (need to understand first how it work)
// it checks the doc existing styles and assign them is assigned to their placeholder automatic
//  Phase 2 – Global Styling Management
    ✅ Create ScriptUI palette/tab with:
    Dropdown for TableStyle (global).
    Dropdown for CellStyle (default for all cells).
    On “Apply”, apply selected TableStyle + CellStyle to the table.
    If no styles exist in document → auto-create “CSV_DefaultTableStyle” + “CSV_DefaultCellStyle”.
//  add info tab : move the first alert about csv to tab in the panel with no if sessions user placeholders add version special thanks
//  remove grey out in hide save space in srtyles and linebreaks
//  ui edits: improve workflow progression ui , ux
//  overview tab fixed ui ux
// fix spacing / make detected field area smart inc or dec depnding on content
//  move script info to advance and be writtern after noorr @2025 | v2.0 | 
//  add fix template analysis and sugguested conf (settings) to take full width
//  rename layout selection to chairpersons
//  move check box next to layout
//  name inline layout congiguratin to inline settings and grid settings
//  move inline settings and grid settings inside chairpersons
//  rename layout sesction to topics
// same as chairpersons
//  in formatting tab title select box checkbox -> the box to write the character then linebreak symbol and remove rest
//  creat new style needs refinment
//  remove profiles
// fix initial panel size
// name change: ClickGenda QuickGend UltraGenda ClickAgenda Autogenda, ProGenda, agendaflow+
//  test all tabs features
//  Use a standardized spacing system (8px increments)
//  code refactoring for the units feature
// copy document setup info
//  remove all unnessacry in drop list 
// --none-- shouldent reset the style it should preserve the elemnt styling but dont add it as a style only
//  option to add paragraph style to the beaked line element. example the first style will be for all the text place holder the second style will be applied for the pragraph after the breakline
// the defult satae for drop list should be with this piririty: if it has style assigned to it, then it should be selected it no style then non is selected
// auto detect the style used
// make the tab smart also hide elements that they arent in the csv with and remove its gap












--------------------------------------------------------------------------------------
// convert to plugin + some identity
// update one page only option
// test preview
// change dependency to excel also and no BOM
// inline confg maybe add text frame option columns with line break
// use windows explorer new ui to choose folder
// plugin export ready to use excel/csv with all data
// auto detect created sessions and updated needed only if text changed, image, chair person added or removed session added or reve
// add ataglange mode


//  Phase 5: Add advanced features (column-specific styles, style preview + breaklines stylings)
    Phase 3 – Per-Column Styling
    🔲 Detect number of CSV columns.
    🔲 Generate a dynamic panel with:
    Column label (“Column 1”, “Column 2”, …).
    Dropdown to assign a CellStyle (with “[Inherit]” as default).
    🔲 On “Apply”, loop through columns and apply the selected style to all cells in that column.
    🔲 Add “+ New Style” button per dropdown to create a new CellStyle if needed.
    Per-Row & Advanced Styling
    🔲 Add row styling rules:
    Header row → specific CellStyle.
//  🔲 Auto-fit & resize options (fit table to text frame, distribute columns evenly).
//  table overseat check box to new page or keep overseats
// cell vetical size automate
// auto fit the text placeholder to fit the max rows in one page
--------------------------------------------------------------------------------------
sessionTitle
sessionNo
sessionTime
chairpersons
topicsTable
topicsTime
topicsTopic
topicsSpeaker
chairAvatar
chairFlag


Section Dropdown Checkboxes text input field panel section scroll list  radio btn
header p style > header Cell style
topics p style > topics Cell style > Topics Table style









🔹 Session-Level Update Logic
Identify sessions in existing document
Each session page has a hidden label like:
page.label = JSON.stringify({id: "S01", title: "Heart Surgery", chairs: [...], topics: [...]})
id could be "Session No" if it’s stable, otherwise "Session Title" + something else.
Parse new CSV into session objects
Example: {id: "S01", title: "Heart Surgery", chairs: [...], topics: [...]}
Compare session lists
If a session exists in both old and new → update it.
If a session exists only in old → remove its page.
If a session exists only in new → create a new page from template.
If overall order changed → reorder session pages to match new CSV order.
🔹 Inside Each Session Page
For each matching session:
1. Text fields (title, time, number)
Compare old.title vs. new.title.
If different → frame.contents = new.title.
(No frame replacement, so formatting stays intact.)
2. Chairpersons
Compare old.chairs array vs. new.chairs array.
If identical → skip.
If different → re-run chair layout routine (inline or grid).
If image placement is enabled → check new chairs, place/update images.
3. Topics
Compare topic lists (old.topics vs. new.topics).
If same number, same order, same text → skip.
If text differs → update text inside cells/frames.
If topics added → insert new row(s) or duplicate topic group template.
If topics removed → delete the row(s)/group(s).
If order changed → move rows/groups to new positions.
4. Images
For each chair/topic with image:
Compare old imageStatus vs. new file existence.
If different (new file appeared, or file missing now) → refresh the link in its frame.
🔹 After Updates
Save updated session metadata back into labels:
Update page.label with new session object.
Example: page.label = JSON.stringify(newSessionData).
Generate an update report (optional but very helpful):
Sessions updated
Sessions added/removed
Topics added/removed/reordered
Images updated/missing
🔹 Pseudocode Summary
for each newSession in newCSV:
    if sessionID exists in document:
        page = findPageBySessionID(sessionID)
        oldSession = JSON.parse(page.label)
        if oldSession.title != newSession.title:
            updateTextFrame(page, "sessionTitle", newSession.title)
        if oldSession.chairs != newSession.chairs:
            updateChairs(page, newSession.chairs)
        compareTopicsAndUpdate(page, oldSession.topics, newSession.topics)
        checkAndUpdateImages(page, newSession.chairs)
        page.label = JSON.stringify(newSession) // refresh metadata
    else:
        createNewSessionPage(newSession)
for each oldSession not in newCSV:
    removePageForSession(oldSession)
reorderPagesToMatchCSV(newCSV.sessionOrder)
This is the logic skeleton.
The “secret sauce” is the metadata labeling (using .label on pages or groups) so the script can remember what was last imported. Without that, the script has no way to know what’s “old” vs. “new.”