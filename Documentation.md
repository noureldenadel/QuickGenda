# QuickGenda v2.0
Complete User Manual & Technical Documentation
Professional InDesign Script for Automated Conference Agenda Generation

## Table of Contents
1. [Introduction](#introduction)
2. [System Requirements & Installation](#system-requirements--installation)
3. [Template Design Guide](#template-design-guide)
4. [CSV Data Preparation](#csv-data-preparation)
5. [Configuration Panel Reference](#configuration-panel-reference)
6. [Styling System](#styling-system)
7. [Image Automation System](#image-automation-system)
8. [Advanced Features](#advanced-features)
9. [Workflow Examples](#workflow-examples)
10. [Troubleshooting Guide](#troubleshooting-guide)
11. [Technical Reference](#technical-reference)
12. [Best Practices](#best-practices)
13. [FAQ](#faq)

## Introduction

QuickGenda is a professional-grade ExtendScript for Adobe InDesign that automates the creation of conference agendas, meeting schedules, and event programs. It transforms CSV data into beautifully formatted InDesign documents with intelligent layout management, automated image placement, comprehensive error reporting, and advanced styling capabilities.

### Key Features

- **Intelligent CSV Processing** - Dynamic column detection with robust parsing
- **Dual Layout Systems** - Table-based and independent row layouts for topics
- **Advanced Grid Management** - Flexible chairperson positioning with centering options
- **Image Automation** - Automatic avatar and flag placement with smart name matching
- **Line Break Processing** - Configurable character-to-linebreak replacement
- **Enhanced Styling System** - Apply paragraph, table, and cell styles directly
- **Style Auto-Detection** - Automatically detect and pre-select styles from templates
- **Inline Style Creation** - Create new styles without leaving the configuration panel
- **Settings Management** - Export/import configurations for workflow consistency
- **Comprehensive Reporting** - Detailed success/failure analysis with overset text detection

### Target Users

- **Conference Organizers** - Academic conferences, professional events
- **Meeting Planners** - Corporate meetings, board sessions
- **InDesign Professionals** - Designers handling multiple agenda projects
- **Administrative Staff** - Personnel managing event documentation
- **Print Shops** - Service providers creating agenda materials

## System Requirements & Installation

### Minimum Requirements

- Adobe InDesign CS6 or later (CC versions recommended)
- Operating System: Windows 7+ or macOS 10.9+
- RAM: 4GB minimum (8GB+ recommended for large datasets)
- Storage: 100MB free space for temporary files
- Template Document: InDesign file with properly labeled placeholders

### Installation Steps

1. Download the QuickGenda.jsx file
2. Place script in one of these locations:
   - User Scripts: [InDesign Install]/Scripts/Scripts Panel/
   - Application Scripts: [User]/AppData/Roaming/Adobe/InDesign/[Version]/Scripts/Scripts Panel/ (Windows)
   - Application Scripts: ~/Library/Preferences/Adobe InDesign/[Version]/Scripts/Scripts Panel/ (macOS)
3. Restart InDesign if currently running
4. Access script via Window → Utilities → Scripts → User (or Application)
5. Test installation by running on sample data

### Verification

Run the script with InDesign open but no documents loaded. You should see:
```
Error: No document is open.
Please open your template file first.
```
This confirms proper installation and script functionality.

## Template Design Guide

### Essential Template Structure

Your InDesign template must contain labeled text frames and objects that serve as placeholders for dynamic content. QuickGenda uses these labels to identify where to place session data, chairperson information, and topic details.

### Required Labels

#### Session Information Labels
- `sessionTitle` - Main session title text frame
- `sessionTime` - Session timing information text frame
- `sessionNo` - Session number/identifier text frame
- `chairpersons` - Chairperson information frame or group prototype

#### Topic Layout Labels

##### For Table-Based Topics:
- `topicsTable` - Table frame for all topic data

##### For Independent Row Topics:
- `topicTime` - Individual topic time text frame
- `topicTitle` - Individual topic title text frame
- `topicSpeaker` - Individual topic speaker text frame

### Template Design Workflow

#### Step 1: Create Base Layout

1. Set up document with desired page size and margins
2. Design session header area with frames for title, time, and session number
3. Create chairperson area for presenter information
4. Design topic section using either table or independent frames
5. Add any decorative elements or branding

#### Step 2: Add Required Labels

1. Select each text frame that will contain dynamic content
2. Open Object → Object Export Options → Alt Text tab
3. Enter the exact label name in the "Alt Text" field, OR
4. Use Window → Utilities → Scripts → Label panel (if available)
5. Verify labels by checking each frame individually

#### Step 3: Configure Chairperson Layout

##### Option A: Inline Layout (Simple)
1. Create single text frame labeled `chairpersons`
2. Size frame to accommodate multiple names
3. Choose comfortable font size and spacing

##### Option B: Grid Layout (Advanced)
1. Create a Group containing:
   - Text frame labeled `chairpersons` (for name)
   - Optional: Image frame labeled `chairAvatar` (for photo)
   - Optional: Image frame labeled `chairFlag` (for flag/country)
2. Position group where first chairperson should appear
3. Ensure group is properly sized and styled

#### Step 4: Configure Topic Layout

##### Option A: Table Layout
1. Insert table frame
2. Label frame as `topicsTable`
3. Set desired width and position
4. Do not pre-format table content (script handles this)

##### Option B: Independent Rows
1. Create three text frames for time, title, and speaker
2. Label them: `topicTime`, `topicTitle`, `topicSpeaker`
3. Position frames as desired for first topic row
4. Optional: Group all three frames together
5. Script will duplicate for additional topics

### Template Testing

1. Save template with meaningful name
2. Create test CSV with 2-3 sessions and multiple topics
3. Run QuickGenda with test data
4. Review output for proper placement and sizing
5. Adjust template as needed and repeat testing
6. Save final template when satisfied with results

### Advanced Template Features

#### Master Pages
- Place consistent elements (headers, footers, logos) on master pages
- Keep dynamic content frames on document pages
- Use master page overrides carefully to maintain script functionality

#### Character and Paragraph Styles
- Apply styles to template frames for consistent formatting
- Script preserves existing character formatting
- Use paragraph styles for consistent topic spacing

#### Color and Effects
- Apply colors, gradients, and effects to template elements
- Script maintains visual formatting while updating content
- Test effects with actual content to ensure readability

## CSV Data Preparation

### File Format Specifications

#### Encoding and Structure
- File Encoding: UTF-8 with or without BOM
- Delimiters: Comma (,) or semicolon (;) - auto-detected
- Text Qualification: Double quotes (") for fields containing delimiters
- Line Endings: Windows (CRLF), Mac (CR), or Unix (LF)
- File Extension: .csv (required)

#### Column Specifications

##### Required Column
| Column Name | Type | Description | Example |
|-------------|------|-------------|---------|
| Session Title | Text | Primary session identifier | "Opening Ceremony" |

##### Optional Columns
| Column Name | Type | Description | Example |
|-------------|------|-------------|---------|
| Session Time | Text | Session timing | "9:00 AM - 10:30 AM" |
| Session No | Text/Number | Session identifier | "1" or "Session A" |
| Chairpersons | Text | Session leaders (use `||` to separate multiple) | "Prof. Dr. John Smith||Dr. Jane Doe" |
| Time | Text | Individual topic timing | "9:00-9:15" |
| Topic Title | Text | Topic description | "Welcome Address" |
| Speaker | Text | Topic presenter | "Prof. Dr. John Smith" |

### Data Structure Principles

#### Session Grouping
QuickGenda groups data by Session Title. When the Session Title changes, a new session begins:

```csv
Session Title,Session Time,Chairpersons,Time,Topic Title,Speaker
Opening Ceremony,9:00 AM,John Smith||Jane Doe,9:00-9:15,Welcome,John Smith
Opening Ceremony,9:00 AM,John Smith||Jane Doe,9:15-9:30,Overview,Jane Doe
Technical Session A,10:30 AM,Robert Johnson,10:30-11:00,Topic 1,Speaker A
Technical Session A,10:30 AM,Robert Johnson,11:00-11:30,Topic 2,Speaker B
```

#### Data Inheritance
- Session-level data (Title, Time, No, Chairpersons) can be repeated or omitted in subsequent rows
- First occurrence of session data is used if later rows are empty
- Topic-level data (Time, Title, Speaker) is row-specific

### Advanced CSV Features

#### Multiple Chairpersons
Use double pipe (||) to separate multiple chairpersons:
```csv
Chairpersons
"Prof. Dr. John Smith||Dr. Jane Doe||Ms. Sarah Wilson"
```

#### Line Break Processing
Include special characters that will be replaced with line breaks:
```csv
Topic Title
"Introduction||Overview||Objectives"
```
Configure in QuickGenda to replace "|" with line breaks.

#### Empty Fields
Empty fields are acceptable and handled gracefully:
```csv
Session Title,Session Time,Topic Title,Speaker
Opening Ceremony,9:00 AM,Welcome,
Opening Ceremony,,Introduction,John Doe
```

### CSV Validation and Testing

#### Pre-Import Checklist
✅ Headers match exactly (case-sensitive)
✅ Required "Session Title" column present
✅ Consistent delimiter usage throughout file
✅ UTF-8 encoding applied
✅ No invalid characters that could break parsing
✅ Chairperson names use || separator

#### Testing Methodology
1. Start small - test with 2-3 sessions first
2. Verify data structure - check grouping works correctly
3. Test special characters - ensure line breaks and formatting work
4. Validate chairperson separation - confirm || delimiter functions
5. Scale up gradually - add more sessions once basics work

### Sample CSV Files

#### Basic Conference Agenda
```csv
Session Title,Session Time,Session No,Chairpersons,Time,Topic Title,Speaker
Opening Ceremony,9:00-10:30 AM,1,Prof. Dr. John Smith||Dr. Jane Doe,9:00-9:15,Welcome Address,Prof. Dr. John Smith
Opening Ceremony,9:00-10:30 AM,1,Prof. Dr. John Smith||Dr. Jane Doe,9:15-9:30,Conference Overview,Dr. Jane Doe
Opening Ceremony,9:00-10:30 AM,1,Prof. Dr. John Smith||Dr. Jane Doe,9:30-10:00,Keynote Speech,Dr. Robert Johnson
Opening Ceremony,9:00-10:30 AM,1,Prof. Dr. John Smith||Dr. Jane Doe,10:00-10:30,Networking Break,
Technical Session A,11:00 AM-12:30 PM,2,Dr. Michael Brown||Prof. Lisa Chen,11:00-11:30,AI in Healthcare,Dr. Sarah Wilson
Technical Session A,11:00 AM-12:30 PM,2,Dr. Michael Brown||Prof. Lisa Chen,11:30-12:00,Machine Learning Applications
Technical Session A,11:00 AM-12:30 PM,2,Dr. Michael Brown||Prof. Lisa Chen,12:00-12:30,Panel Discussion,Multiple Speakers
```

#### Corporate Meeting Agenda
```csv
Session Title,Session Time,Chairpersons,Time,Topic Title,Speaker
Board Meeting,2:00-4:00 PM,John Smith|CEO||Jane Doe|CFO,2:00-2:15,Financial Report,Jane Doe
Board Meeting,2:00-4:00 PM,John Smith|CEO||Jane Doe|CFO,2:15-2:45,Strategic Planning,John Smith
Board Meeting,2:00-4:00 PM,John Smith|CEO||Jane Doe|CFO,2:45-3:00,Q3 Review,Robert Johnson
```

## Configuration Panel Reference

The QuickGenda Configuration panel provides comprehensive control over how your agenda is generated. It consists of four main tabs: Chairpersons, Topics, Line Breaks, and Styling.

### Chairpersons Tab

#### Layout Selection
- **Inline (single frame)**
  - Places all chairpersons in the labeled `chairpersons` text frame
  - Suitable for simple layouts with few chairpersons
  - Minimal template setup required
  
- **Separate frames grid**
  - Creates individual frames for each chairperson
  - Enables image automation for photos and flags
  - Requires group-based template setup
  - Provides precise positioning control

#### Inline Layout Configuration
- **Comma separated**
  - Joins chairpersons with commas: "John Smith, Jane Doe, Robert Johnson"
  - Standard format for most applications
  - Compact presentation
  
- **Line breaks**
  - Places each chairperson on separate line
  - Better readability for longer names
  - Useful when space allows vertical layout

#### Grid Layout Configuration
- **Fill order**
  - Row-first: Fills horizontally before moving to next row
  - Column-first: Fills vertically before moving to next column

- **Columns/Rows**
  - Row-first mode: Set number of columns, rows calculated automatically
  - Column-first mode: Set number of rows, columns calculated automatically
  - Valid range: 1-99

- **Spacing**
  - Column spacing: Horizontal distance between chairperson frames
  - Row spacing: Vertical distance between chairperson frames
  - Units: Points (pt), millimeters (mm), centimeters (cm), pixels (px)

- **Center grid horizontally**
  - Centers the entire chairperson grid horizontally on the page
  - Maintains original vertical position
  - Centers based on actual number of chairpersons, not full grid

#### Image Automation
- **Enable automatic image placement**
  - Activates intelligent image placement system
  - Requires labeled image frames within chairperson groups
  - Searches for avatar and flag images automatically

- **Image folder**
  - Path to folder containing chairperson images
  - Supports network paths and local directories
  - Browse button provides folder selection dialog

- **Image fitting**
  - Use Template Default: Maintains existing frame fitting settings
  - Fill Frame: Scales image to fill frame completely (may crop)
  - Fit Proportionally: Scales to fit while maintaining aspect ratio
  - Fit Content to Frame: Scales to exact frame dimensions
  - Center Content: Centers image in frame without scaling

### Topics Tab

#### Layout Mode Selection
- **Table Layout**
  - Creates formatted table with rows and columns
  - Suitable for structured data presentation
  - Requires `topicsTable` labeled frame in template
  
- **Independent Row Layout**
  - Creates separate frames for each topic component
  - Allows custom positioning and styling
  - Requires `topicTime`, `topicTitle`, `topicSpeaker` labeled frames

#### Table Layout Configuration
- **Include Table Header Row**
  - Adds header row with "Time", "Topic", "Speaker" labels
  - Improves data readability and structure
  - Uses InDesign's built-in header row formatting

#### Independent Row Configuration
- **Vertical spacing**
  - Distance between topic rows
  - Applies to group spacing if topics are grouped
  - Units: Points (pt), millimeters (mm), centimeters (cm), pixels (px)

### Line Breaks Tab

The Line Breaks tab allows you to replace specific characters with line breaks in any field, enabling better text formatting and presentation.

#### Session Fields
- **Session Title**
  - Replaces specified character with line breaks in session titles
  - Useful for multi-line session names
  - Default character: "|" (pipe)

- **Session Time**
  - Processes time information for line breaks
  - Enables multi-line time displays
  - Common use: separating date and time

- **Session No**
  - Applies line break processing to session numbers
  - Useful for complex session identifiers

#### Chairpersons
- **Chairpersons**
  - Processes chairperson names for line breaks
  - Applies to individual chairperson names, not the separator
  - Useful for titles and affiliations: "Prof. Dr. John Smith|University of Technology"

#### Topics
- **Topic Time**
  - Processes topic timing information
  - Enables multi-line time displays
  - Example: "10:00 AM|Room A"

- **Topic Title**
  - Processes topic titles for line breaks
  - Enables multi-line topic descriptions
  - Common use: "Main Topic|Subtitle|Additional Info"

- **Topic Speaker**
  - Processes speaker information
  - Enables multi-line speaker details
  - Example: "Dr. John Smith|University Name|Department"

### Styling Tab

The Styling tab provides advanced control over document formatting through InDesign's paragraph, table, and cell styles.

#### Field Styling
- **Session Title Style**
  - Apply predefined paragraph style to session titles
  - Supports line break processing
  - Auto-detection from template document

- **Session Time Style**
  - Apply predefined paragraph style to session times
  - Supports line break processing

- **Session Number Style**
  - Apply predefined paragraph style to session numbers
  - Supports line break processing

- **Chairpersons Style**
  - Apply predefined paragraph style to chairperson names
  - Supports line break processing

- **Topic Time Style**
  - Apply predefined paragraph style to topic times
  - Supports line break processing

- **Topic Title Style**
  - Apply predefined paragraph style to topic titles
  - Supports line break processing

- **Topic Speaker Style**
  - Apply predefined paragraph style to topic speakers
  - Supports line break processing

#### Table Styling
- **Table Style**
  - Apply predefined table style to topics table layout
  - Auto-detection from template document

- **Cell Style**
  - Apply predefined cell style to topics table layout
  - Auto-detection from template document

- **Table Paragraph Style**
  - Apply predefined paragraph style to table content
  - Supports line break processing

#### Style Creation
- **Create New Style**
  - Create new paragraph, table, or cell styles directly from the panel
  - Instantly available in all dropdowns
  - No need to leave configuration panel

### Settings Management

#### Import All Settings
- Loads previously saved configuration from JSON file
- Applies settings to all tabs simultaneously
- Preserves workflow consistency across projects

#### Export All Settings
- Saves current configuration to JSON file
- Includes all tab settings in single file
- Enables sharing configurations between users

## Styling System

QuickGenda v2.0 introduces a powerful styling system that allows you to apply professional formatting to your agendas directly from the configuration panel.

### Paragraph Styles

Paragraph styles control the formatting of text elements including:
- Session titles, times, and numbers
- Chairperson names
- Topic times, titles, and speakers

#### Applying Paragraph Styles
1. Select a paragraph style from the dropdown in the Styling tab
2. The style will be applied to the corresponding text elements
3. Existing character formatting is preserved
4. Styles are applied during agenda generation

#### Style Auto-Detection
QuickGenda automatically detects paragraph styles from your template document:
- Scans the active document for existing paragraph styles
- Pre-selects matching styles when available
- Shows "— None —" when no matching style is found

### Table Styles

Table styles control the formatting of the topics table including:
- Table borders and shading
- Header row formatting
- Alternate row colors

#### Applying Table Styles
1. Select a table style from the dropdown in the Styling tab
2. The style will be applied to the topics table
3. Works only with table layout mode
4. Styles are applied during agenda generation

### Cell Styles

Cell styles control the formatting of individual table cells including:
- Cell borders and shading
- Text inset spacing
- Cell background colors

#### Applying Cell Styles
1. Select a cell style from the dropdown in the Styling tab
2. The style will be applied to all cells in the topics table
3. Works only with table layout mode
4. Styles are applied during agenda generation

### Inline Style Creation

QuickGenda allows you to create new styles without leaving the configuration panel:

1. Navigate to the "Create New Style" section in the Styling tab
2. Enter a name for your new style
3. Select the style type (Paragraph, Table, or Cell)
4. Click the "+ Create" button
5. The new style will immediately appear in all relevant dropdowns

### Best Practices for Styling

1. **Prepare Styles in Template**
   - Create paragraph, table, and cell styles in your template document
   - Use descriptive names that match their purpose
   - Apply base formatting before running QuickGenda

2. **Use Style Auto-Detection**
   - Name your styles to match QuickGenda's field names
   - Take advantage of automatic style detection
   - Verify that detected styles are correct

3. **Test Style Application**
   - Run QuickGenda with sample data first
   - Check that styles are applied correctly
   - Adjust styles in template if needed

## Image Automation System

The Image Automation System is one of QuickGenda's most advanced features, providing intelligent placement of chairperson photos and flags with sophisticated name matching and file detection.

### System Overview

#### Capabilities
- Automatic image detection based on chairperson names
- Multiple file format support (JPG, PNG, TIF, PSD)
- Intelligent name matching with variation generation
- Avatar and flag placement in designated frames
- Title stripping from academic and professional names
- Comprehensive error reporting with search attempt logging

#### Requirements
- Grid layout mode must be selected for chairpersons
- Group-based template with labeled image frames
- Organized image folder with consistent naming
- Enable automatic image placement checkbox activated

### Template Setup for Images

#### Frame Labeling Requirements
Within each chairperson group, create image frames with these exact labels:
- `chairAvatar` - For chairperson photo/headshot
- `chairFlag` - For country flag or organizational logo

#### Template Structure Example
```
Group: Chairperson Prototype
├──── Text Frame (label: "chairpersons") - for name
├──── Image Frame (label: "chairAvatar") - for photo
└──── Image Frame (label: "chairFlag") - for flag
```

#### Frame Positioning
- Position frames where images should appear
- Size frames appropriately for image content
- Apply any desired stroke, fill, or effects
- Test with sample images to verify appearance

### Image Folder Organization

#### Folder Structure
Create a dedicated folder containing all chairperson images:
```
📁 Chairperson_Images/
📸 Avatar Images:
├──── john smith.jpg
├──── jane doe.png
├──── robert johnson.tiff
└──── sarah wilson.psd
🏳️ Flag Images:
├──── flag-john smith.jpg
├──── jane doe-flag.png
├──── flag_robert_johnson.jpg
└──── flag-sarah-wilson.jpg
```

#### Supported Formats
- JPEG/JPG - Most common, good compression
- PNG - Supports transparency, good for flags
- TIFF/TIF - High quality, large file size
- PSD - Photoshop format, maintains layers

### Name Processing Intelligence

#### Title Stripping
QuickGenda automatically removes academic and professional titles:
| Original Name | Processed Name |
|---------------|----------------|
| Prof. Dr. John Smith | John Smith |
| Dr. Jane Doe, MD | Jane Doe |
| Associate Prof. Robert Johnson | Robert Johnson |
| Mr. David Wilson | David Wilson |
| Prof. Dr. Med. Sarah Chen | Sarah Chen |

#### Supported Title Patterns
- Academic: Prof., Dr., Associate Prof., Assoc. Prof.
- Medical: MD, PhD, M.D., Ph.D.
- Courtesy: Mr., Mrs., Ms., Miss
- Combined: Prof. Dr. Med., Prof. Dr., Dr. Med.

### Name Variation Generation

#### Avatar Image Variations
For chairperson "John Smith", the system searches for:

**Case Variations:**
- john smith.jpg
- John Smith.jpg
- JOHN SMITH.jpg
- johnsmith.jpg

**Separator Variations:**
- john_smith.jpg
- john-smith.jpg
- john smith.jpg

**Format Variations:** Each name variation is tried with:
- .jpg / .jpeg
- .png
- .tif / .tiff
- .psd

#### Flag Image Variations
For chairperson "John Smith", the system searches for flag images:

**Prefix Patterns:**
- flag-john smith.jpg
- flag_john smith.jpg
- flag john smith.jpg
- flagjohnsmith.jpg

**Suffix Patterns:**
- john smith-flag.jpg
- john smith_flag.jpg
- john smith flag.jpg
- johnsmithflag.jpg

**Case and Format Combinations:** Each pattern is tried with all case and format variations.

### Image Fitting Options

#### Use Template Default
- Maintains existing frame fitting settings
- Preserves designer's intended appearance
- Recommended for pre-configured templates

#### Fill Frame
- Scales image proportionally to fill entire frame
- May crop parts of image
- Good for consistent frame filling

#### Fit Proportionally
- Scales image to fit within frame bounds
- Maintains original aspect ratio
- Prevents distortion, may leave empty space

#### Fit Content to Frame
- Scales image to exact frame dimensions
- May distort image if aspect ratios differ
- Ensures complete frame coverage

#### Center Content
- Centers image within frame without scaling
- Maintains original image size
- May result in cropping or empty space

### Search Process Details

#### Search Algorithm
1. Extract clean name from CSV data (remove titles)
2. Generate name variations (case, spacing, separators)
3. Search for avatar image using all variations and formats
4. If avatar found, search for flag using flag patterns
5. Place images in labeled frames with specified fitting
6. Log all search attempts for reporting

#### Search Order Priority
1. Exact case match with original spacing
2. Title case variations
3. Lowercase variations
4. No spaces variations
5. Underscore separator variations
6. Hyphen separator variations

### Error Handling and Reporting

#### Success Tracking
- Successful placements logged with filenames
- Image fitting applied according to settings
- File paths recorded for verification

#### Failure Analysis
- Missing images logged with clean names
- Search attempts recorded (up to first 10 variations)
- Folder existence verified
- File format attempts documented

#### Report Output Example
```
IMAGE PLACEMENT REPORT:
============================================
Image folder: C:\Projects\Conference\Images
Total chairpersons processed: 15
Successfully placed: 12/15 images
Missing images: 3/15

SUCCESSFUL PLACEMENTS:
[[SUCCESS]] John Smith -> john__smith.jpg + flag--john--smith.png
[[SUCCESS]] Jane Doe -> jane doe.jpg + jane__doe--flag.jpg

MISSING IMAGES:
[[MISSING]] Robert Johnson
Searched for: robert johnson.jpg, robert__johnson.jpg, robert--johnson.jpg...
```

## Advanced Features

### Template Inheritance System

QuickGenda includes sophisticated template management that preserves document settings and structure across generated agendas.

#### Document Settings Preservation
- Page dimensions and orientation
- Margin settings (top, bottom, left, right)
- Facing pages configuration
- Page binding and direction
- Bleed and slug settings
- Master page relationships

#### Template Duplication Process
1. Open template document in InDesign
2. Create new document with inherited settings
3. Duplicate template pages to new document
4. Remove original blank page
5. Use first page as master template
6. Duplicate master for each session
7. Remove master after processing complete

### Dynamic Layout Detection

#### Automatic Capability Detection
QuickGenda analyzes your template and automatically determines available layout options:

**Independent Topic Layout Detection:**
- Searches for `topicTime`, `topicTitle`, `topicSpeaker` frames
- Enables independent row layout option if all found
- Disables option if any frames missing

**Table Topic Layout Detection:**
- Searches for `topicsTable` frame
- Enables table layout option if found
- Sets as default if independent layout unavailable

#### Layout Option Management
- Auto-enable appropriate layout based on template
- Disable incompatible options in configuration panel
- Provide user feedback about detected capabilities
- Fallback gracefully if preferred option unavailable

### Intelligent Grid Centering

#### Horizontal Centering Algorithm
Traditional centering centers the full grid space, but QuickGenda centers based on actual content:

**Standard Centering Problem:**
```
Grid space: [ X X X ]
Chairpersons: Only 2 people
Result: [X X ] (off-center)
```

**QuickGenda Smart Centering:**
```
Grid space: [ X X X ]
Chairpersons: Only 2 people
Result: [ X X ] (properly centered)
```

#### Centering Process
1. Calculate actual chairperson count
2. Determine rows/columns needed for actual count
3. Calculate space required for actual layout
4. Center the used space horizontally on page
5. Maintain original vertical position

#### Row-First vs Column-First Centering
- Row-first: Centers based on widest row
- Column-first: Centers based on actual columns used
- Handles partial rows intelligently
- Accounts for spacing in calculations

### Advanced CSV Processing

#### Robust Parsing Engine
- Multi-delimiter support (comma, semicolon auto-detection)
- Quoted field handling with escape sequence support
- UTF-8 BOM detection and removal
- Line ending normalization (Windows, Mac, Unix)
- Empty line skipping and data validation

#### Dynamic Field Detection
- Session grouping logic based on Session Title changes
- Data inheritance within sessions
- Topic aggregation per session
- Empty field handling gracefully
- Duplicate detection and management

### Memory Management

#### Efficient Object Handling
- Cleanup duplicate objects after each session
- Remove prototype elements when no longer needed
- Minimize DOM manipulation for performance
- Batch operations where possible
- Progressive garbage collection

#### Clone Management System
```javascript
// Systematic cleanup approach
removeExistingChairClone(page); // Before creating new
removeExistingTopicClone(page); // Before creating new
```

### Error Recovery Systems

#### Graceful Degradation
- Continue processing when non-critical errors occur
- Skip problematic records rather than failing completely
- Provide detailed error reporting for troubleshooting
- Maintain partial functionality when features unavailable

#### Validation Checkpoints
- Template validation before processing
- CSV structure validation after parsing
- Image folder validation before automation
- Frame existence verification before content placement

## Workflow Examples

### Example 1: Academic Conference Agenda

#### Scenario
Large academic conference with:
- 15 sessions over 3 days
- Multiple chairpersons per session
- 3-5 topics per session
- International speakers (need flag images)
- Professional headshots available

#### Template Setup
```
Session Header Area:
├──── sessionTitle (large, bold text frame)
├──── sessionTime (medium text frame)
└──── sessionNo (small text frame)

Chairperson Area (Grid Layout):
└──── Group: Chairperson Prototype
     ├──── chairpersons (text frame for name)
     ├──── chairAvatar (square image frame)
     └──── chairFlag (small rectangular image frame)

Topic Area (Table Layout):
└──── topicsTable (wide table frame)
```

#### CSV Structure
```csv
Session Title,Session Time,Session No,Chairpersons,Time,Topic Title,Speaker
Opening Plenary,Day 1 -- 9:00 AM,S001,Prof. Dr. Sarah Chen||Dr. Marcus Weber,9:00-9:15,Welcome Address,Prof. Dr. Sarah
Opening Plenary,Day 1 -- 9:00 AM,S001,Prof. Dr. Sarah Chen||Dr. Marcus Weber,9:15-9:45,Keynote: Future of AI,Dr. Elena
Technical Session A,Day 1 -- 10:30 AM,S002,Dr. James Liu||Prof. Maria Santos,10:30-11:00,Machine Learning in Healthcare
Technical Session A,Day 1 -- 10:30 AM,S002,Dr. James Liu||Prof. Maria Santos,11:00-11:30,Neural Networks Applications,P
```

#### Configuration Settings
```
Chairpersons Tab:
-- Layout: Separate frames grid
-- Fill order: Row-first
-- Columns: 2
-- Column spacing: 15 pt
-- Center grid horizontally: ✓
-- Enable image automation: ✓
-- Image folder: //Conference2025/Photos
-- Image fitting: Fit Proportionally

Topics Tab:
-- Layout: Table Layout
-- Include header row: ✓

Line Breaks Tab:
-- Speaker: Replace "||" with line breaks (for titles/affiliations)

Styling Tab:
-- Session Title: Heading 1
-- Session Time: Body Text
-- Session No: Small Text
-- Chairpersons: Person Name
-- Topic Time: Table Cell
-- Topic Title: Table Cell Bold
-- Topic Speaker: Table Cell
-- Table Style: Agenda Table
-- Cell Style: Agenda Cell
```

#### Expected Output
- Each session on separate page
- Chairpersons in 2-column grid with photos and flags
- Topics in formatted table with headers
- Professional appearance with consistent spacing

### Example 2: Corporate Board Meeting

#### Scenario
Quarterly board meeting with:
- 3 main sessions
- 2-3 board members per session
- Detailed agenda items
- Corporate headshots in company database

#### Template Setup
```
Corporate Header:
├──── Company logo (fixed element)
├──── sessionTitle (corporate font)
└──── sessionTime (smaller, aligned right)

Executive Area (Inline Layout):
└──── chairpersons (single text frame, comma-separated)

Agenda Items (Independent Rows):
├──── topicTime (left column, narrow)
├──── topicTitle (center column, wide)
└──── topicSpeaker (right column, medium)
```

#### CSV Structure
```csv
Session Title,Session Time,Chairpersons,Time,Topic Title,Speaker
Financial Review,2:00-3:30 PM,John Smith|CEO||Sarah Johnson|CFO||Robert Davis|COO,2:00-2:15,Q3 Financial Results,Sa
Financial Review,2:00-3:30 PM,John Smith|CEO||Sarah Johnson|CFO||Robert Davis|COO,2:15-2:45,Budget Planning 2026,
Strategic Planning,4:00-5:30 PM,John Smith|CEO||Michael Brown|CTO,4:00-4:30,Technology Roadmap,Michael Brown
Strategic Planning,4:00-5:30 PM,John Smith|CEO||Michael Brown|CTO,4:30-5:00,Market Expansion,John Smith
```

#### Configuration Settings
```
Chairpersons Tab:
-- Layout: Inline (single frame)
-- Separator: Comma separated

Topics Tab:
-- Layout: Independent Row Layout
-- Vertical spacing: 8 pt

Line Breaks Tab:
-- Chairpersons: Replace "||" with line breaks (for titles)
-- Topic Title: Replace "||" with line breaks (for sub-items)

Styling Tab:
-- Session Title: Corporate Heading
-- Session Time: Small Text
-- Chairpersons: Executive Name
-- Topic Time: Agenda Time
-- Topic Title: Agenda Item
-- Topic Speaker: Presenter Name
```

### Example 3: Workshop Schedule

#### Scenario
Educational workshop with:
- Multiple parallel sessions
- Room assignments in time field
- Presenter photos needed
- Simple, clean layout

#### Template Setup
```
Workshop Header:
├──── sessionTitle (workshop name)
├──── sessionTime (date/time range)
└──── sessionNo (workshop code)

Facilitator Area (Grid with Images):
└──── Group: Facilitator
     ├──── chairpersons (name and credentials)
     └──── chairAvatar (professional headshot)

Schedule (Table with Room Info):
└──── topicsTable (includes room assignment in time column)
```

#### CSV Structure
```csv
Session Title,Session Time,Session No,Chairpersons,Time,Topic Title,Speaker
Digital Marketing Fundamentals,March 15 -- 9:00 AM,WS101,Dr. Lisa Chen||Mark Rodriguez,9:00-10:30||Room A101,Intro
Digital Marketing Fundamentals,March 15 -- 9:00 AM,WS101,Dr. Lisa Chen||Mark Rodriguez,10:
```

## Troubleshooting Guide

### Common Issues

#### "No sessions found in CSV"
- Check that you have a "Session Title" column
- Verify CSV encoding is UTF-8
- Ensure headers match exactly (case-sensitive)

#### "Could not find any placeholders"
- Verify text frame labels match required names
- Check for typos in frame labels
- Ensure labels are on text frames, not groups

#### "Overset text detected"
- Increase text frame sizes in template
- Reduce font size or content
- Check line spacing settings

#### Images not placing
- Verify image folder path is correct
- Check image file names match chairperson names
- Ensure image frames are labeled `chairAvatar` and `chairFlag`
- Confirm image frames are within chairperson groups

### Error Reporting

QuickGenda provides comprehensive error reporting to help diagnose issues:

#### CSV Parsing Errors
- Line-by-line parsing failures
- Invalid character detection
- Delimiter mismatch warnings

#### Template Validation Errors
- Missing required labels
- Incorrect frame types
- Layout capability mismatches

#### Image Processing Errors
- File not found reports
- Format compatibility issues
- Name matching failures

### Performance Optimization

#### Large Dataset Handling
- Process data in smaller batches
- Optimize template with efficient styles
- Close unnecessary applications during processing

#### Memory Management
- Clear clipboard before running script
- Close unused documents
- Restart InDesign periodically for long sessions

## Technical Reference

### Script Architecture

QuickGenda follows a modular architecture with clearly defined components:

#### Main Processing Flow
1. Template analysis and validation
2. CSV parsing and data structure creation
3. User configuration collection
4. Document generation and formatting
5. Reporting and cleanup

#### Core Modules
- **CSV Parser** - Handles dynamic column detection and data grouping
- **Template Analyzer** - Detects layout capabilities and frame labels
- **Configuration UI** - Unified settings panel with tabbed interface
- **Document Generator** - Creates and populates agenda pages
- **Image Processor** - Handles avatar and flag placement
- **Style Manager** - Applies paragraph, table, and cell styles
- **Reporting System** - Generates detailed processing reports

### API Reference

#### Core Functions
- `main()` - Entry point for script execution
- `parseCSVDynamic()` - Parses CSV with automatic delimiter detection
- `getUnifiedSettingsPanel()` - Creates and manages configuration UI
- `processSessions()` - Generates agenda pages from session data
- `applyStyling()` - Applies document styles to content elements

#### Utility Functions
- `findItemByLabel()` - Locates frames by label in template
- `detectIndependentTopicSetup()` - Checks for independent topic layout
- `copyDocumentSettings()` - Preserves template document settings
- `createParagraphStyleIfMissing()` - Creates styles when needed

### Supported InDesign Versions

QuickGenda is compatible with:
- InDesign CS6
- InDesign CC 2014-2025
- InDesign 2025 and later versions

### File Format Support

#### Input Formats
- CSV (UTF-8 encoded)
- Comma or semicolon delimited
- Quoted fields with escape sequences

#### Image Formats
- JPEG/JPG
- PNG
- TIFF/TIF
- PSD

#### Export Formats
- InDesign document (.indd)
- PDF (through InDesign's export functionality)
- Other formats via InDesign's export options

## Best Practices

### Template Design

#### Consistency
- Use consistent frame sizes for grid layouts
- Maintain uniform spacing throughout document
- Apply consistent styling to similar elements

#### Flexibility
- Leave sufficient space for content expansion
- Design for variable content lengths
- Use master pages for repeated elements

#### Testing
- Test with sample data before production
- Verify all layout options work correctly
- Check edge cases with minimal/maximal data

### CSV Preparation

#### Data Quality
- Clean data before import (remove extra spaces)
- Use consistent naming for chairpersons
- Include all required columns even if some are empty
- Validate data structure before processing

#### Formatting
- Use UTF-8 encoding for international characters
- Apply consistent date/time formatting
- Standardize academic titles and credentials

#### Organization
- Keep CSV files organized in project folders
- Maintain backup copies of original data
- Document CSV structure for future reference

### Image Organization

#### File Management
- Use consistent naming conventions
- Keep images in dedicated folder
- Maintain organized directory structure

#### Quality Control
- Use high-quality images (300+ DPI)
- Maintain consistent aspect ratios
- Optimize file sizes for performance

#### Backup and Versioning
- Keep original high-resolution images
- Track changes to image collections
- Document image sources for attribution

## FAQ

### General Questions

#### What is QuickGenda?
QuickGenda is a professional InDesign script that automates the creation of conference agendas, meeting schedules, and event programs from CSV data.

#### What are the system requirements?
QuickGenda requires Adobe InDesign CS6 or later and a template document with properly labeled placeholders.

#### Is QuickGenda free to use?
QuickGenda is available for free use, but may require proper attribution. Check the license file for specific terms.

### Template Design

#### How do I label frames in InDesign?
Select the frame, go to Object → Object Export Options → Alt Text, and enter the label in the "Alt Text" field.

#### Can I use master pages with QuickGenda?
Yes, QuickGenda preserves master page elements and relationships during document generation.

#### What happens if I don't have all the required labels?
QuickGenda will disable features that require missing labels and inform you of any limitations.

### CSV Data

#### What CSV format does QuickGenda expect?
QuickGenda supports UTF-8 encoded CSV files with comma or semicolon delimiters.

#### How do I separate multiple chairpersons?
Use double pipe (||) to separate multiple chairpersons in the Chairpersons column.

#### Can I have empty fields in my CSV?
Yes, empty fields are handled gracefully and will result in empty content in the generated document.

### Styling

#### How do I apply styles to my agenda?
Use the Styling tab in the configuration panel to select paragraph, table, and cell styles for different elements.

#### Can I create new styles from within QuickGenda?
Yes, use the "Create New Style" section in the Styling tab to create paragraph, table, or cell styles.

#### What if my template doesn't have the styles I want to use?
QuickGenda will show "— None —" for missing styles. You can create new styles or modify your template.

### Image Automation

#### What image formats are supported?
QuickGenda supports JPG, PNG, TIFF, and PSD formats.

#### How do I organize my image folder?
Create a dedicated folder with avatar images and flag images following the naming conventions in the documentation.

#### What if an image isn't found?
QuickGenda will log missing images in the report and continue processing other content.

### Troubleshooting

#### Why am I getting "No document is open" error?
This error appears when you run QuickGenda without having a template document open in InDesign.

#### Why aren't my images appearing?
Check that image automation is enabled, the folder path is correct, and image names match chairperson names.

#### How do I fix overset text issues?
Increase the size of text frames in your template or reduce the amount of content.

---

**QuickGenda v2.0** - Transform your conference planning workflow with intelligent automation.