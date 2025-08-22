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
//  option to add optional style to beaklines and if there is no styles assigned it will generate one for each placeholder with name like the csv placeholder
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






Section Dropdown Checkboxes text input field panel section scroll list  radio btn
header p style > header Cell style
topics p style > topics Cell style > Topics Table style









--------------------------------------------------------------------------------------
// convert to plugin + some identity
// update one page only option
// test preview
// change dependency to excel also and no BOM
// inline confg maybe add text frame option columns with line break
// use windows explorer new ui to choose folder
// plugin export ready to use excel/csv with all data

// add ataglange mode
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