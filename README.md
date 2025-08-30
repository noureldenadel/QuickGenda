# QuickGenda v2.0

**Professional InDesign Script for Automated Conference Agenda Generation**

![quickgenda](https://github.com/user-attachments/assets/514be2e0-8c67-4a7b-b735-9117122d3e7f)

Transform your CSV data into beautifully formatted InDesign agendas with intelligent layout management, automated image placement, and comprehensive error reporting.

## 🚀 Quick Start

1. **Prepare your InDesign template** with labeled placeholders
2. **Format your CSV data** according to the specifications below
3. **Run the script** in InDesign with your template open
4. **Configure layout and styling options** in the QuickGenda Configuration panel
5. **Import your CSV** and let QuickGenda do the rest!

## 📋 System Requirements

- **InDesign CS6 or later**
- **Template document** with properly labeled placeholders
- **CSV file** with session and topic data
- **Optional:** Image folder for chairperson photos/flags

## 📊 CSV Format Requirements

### Required Column
- `Session Title` - Primary session identifier *(required)*

### Optional Columns
- `Session Time` - Session timing information
- `Session No` - Session numbering/identification  
- `Chairpersons` - Session leaders (use `||` to separate multiple chairpersons)
- `Time` - Individual topic timing
- `Topic Title` - Topic descriptions
- `Speaker` - Topic presenter information

### CSV Example
```csv
Session Title,Session Time,Session No,Chairpersons,Time,Topic Title,Speaker
Opening Ceremony,9:00 AM,1,Prof. Dr. John Smith||Dr. Jane Doe,9:00-9:15,Welcome Address,Prof. Dr. John Smith
Opening Ceremony,9:00 AM,1,Prof. Dr. John Smith||Dr. Jane Doe,9:15-9:30,Opening Remarks,Dr. Jane Doe
Technical Session A,10:30 AM,2,Dr. Robert Johnson||Ms. Sarah Wilson,10:30-11:00,AI in Healthcare,Dr. Michael Brown
Technical Session A,10:30 AM,2,Dr. Robert Johnson||Ms. Sarah Wilson,11:00-11:30,Machine Learning Trends,Prof. Lisa Chen
```

### Important CSV Notes
- ✅ **UTF-8 encoding** recommended
- ✅ **Comma or semicolon** delimited (auto-detected)
- ✅ **Double pipe `||`** separates multiple chairpersons
- ✅ **Headers must match exactly** (case-sensitive)
- ✅ **Empty fields** are acceptable

## 🏷️ Template Label Requirements

### Essential Labels (place on text frames)
- `sessionTitle` - Session title text frame
- `sessionTime` - Session time text frame  
- `sessionNo` - Session number text frame
- `chairpersons` - Chairpersons text frame or group prototype

### Topic Layout Labels

#### For Table Layout
- `topicsTable` - Table frame for topic data

#### For Independent Row Layout
- `topicTime` - Topic time text frame
- `topicTitle` - Topic title text frame
- `topicSpeaker` - Topic speaker text frame

### Image Automation Labels (Optional)
*Place these labels on image frames within chairperson groups:*
- `chairAvatar` - Avatar/photo image frame
- `chairFlag` - Flag/country image frame

## 🖼️ Image Automation Setup

### Folder Structure
Create a folder with chairperson images using these naming conventions:

```
📁 Chairperson_Images/
   📄 john smith.jpg          ← Avatar image
   📄 flag-john smith.jpg     ← Flag image
   📄 jane doe.png            ← Avatar image  
   📄 flag_jane_doe.png       ← Flag image
   📄 robert johnson.tiff     ← Avatar image
   📄 flag-robert-johnson.jpg ← Flag image
```

### Supported Image Formats
- JPG/JPEG
- PNG  
- TIF/TIFF
- PSD

### Image Naming Patterns
The script automatically tries multiple naming variations:

**Avatar Images:**
- `john smith.jpg`
- `johnsmith.jpg`
- `john_smith.jpg`
- `john-smith.jpg`
- `John Smith.jpg`

**Flag Images:**
- `flag-john smith.jpg`
- `john smith-flag.jpg`
- `flag_john_smith.jpg`
- `john_smith_flag.jpg`
- `flagjohnsmith.jpg`

### Name Processing
QuickGenda automatically strips academic titles:
- `Prof. Dr. John Smith|CEO` → `John Smith`
- `Dr. Jane Doe, MD` → `Jane Doe`
- `Associate Prof. Robert Johnson` → `Robert Johnson`

## ⚙️ Configuration Options

### Chairpersons Layout
- **Inline Mode:** Text in single frame (comma-separated or line breaks)
- **Grid Mode:** Individual frames in customizable grid layout
  - Configurable rows/columns
  - Adjustable spacing
  - Horizontal centering option
  - Image automation integration

### Topics Layout  
- **Table Layout:** Topics in formatted table with headers
- **Independent Rows:** Individual topic frames with custom spacing

### Line Breaks
- Replace any character with line breaks in any field
- Apply to session info, chairpersons, or topic data
- Configurable per field type

### Styling Options
- **Paragraph Styles:** Apply predefined paragraph styles to session titles, times, numbers, chairpersons, and topic fields
- **Table Styles:** Apply table styles to topics table layout
- **Cell Styles:** Apply cell styles to topics table layout
- **Style Creation:** Create new paragraph, table, or cell styles directly from the configuration panel
- **Auto-Detection:** Automatically detect and pre-select styles from your template document

## 🚀 Usage Instructions

1. **Open your template** in InDesign with proper labels
2. **Run the script:** File → Scripts → Scripts Panel → QuickGenda.jsx
3. **Select your CSV file** when prompted
4. **Review detected fields** in the confirmation dialog
5. **Configure layout and styling options** in the QuickGenda Configuration panel:
   - Set chairperson layout (inline/grid)
   - Choose topic layout (table/independent)  
   - Configure line break replacements
   - Apply paragraph, table, and cell styles
   - Create new styles directly in the panel
   - Enable image automation if desired
6. **Click OK** to generate your agenda
7. **Review the report** for any issues or missing images

## 📊 Reports & Feedback

QuickGenda generates comprehensive reports including:
- ✅ **Import success summary** - sessions and topics processed
- ⚠️ **Overset text alerts** - text that doesn't fit in frames
- 🖼️ **Image placement results** - successful and failed image placements
- 📝 **Missing image details** - what files were searched for
- 🎨 **Applied styles report** - paragraph, table, and cell styles used
- 🔍 **Troubleshooting info** - detailed error information

## 🛠️ Troubleshooting

### Common Issues

**"No sessions found in CSV"**
- Check that you have a "Session Title" column
- Verify CSV encoding is UTF-8
- Ensure headers match exactly (case-sensitive)

**"Could not find any placeholders"**
- Verify text frame labels match required names
- Check for typos in frame labels
- Ensure labels are on text frames, not groups

**"Overset text detected"**
- Increase text frame sizes in template
- Reduce font size or content
- Check line spacing settings

**Images not placing**
- Verify image folder path is correct
- Check image file names match chairperson names
- Ensure image frames are labeled `chairAvatar` and `chairFlag`
- Confirm image frames are within chairperson groups

## 💾 Settings Management

- **Export Settings:** Save your configuration for reuse
- **Import Settings:** Load previously saved configurations
- **JSON Format:** Settings stored in standard JSON format
- **Cross-Session:** Settings persist across InDesign sessions

## 🎯 Best Practices

### Template Design
- Use consistent frame sizes for grid layouts
- Leave sufficient space for content expansion
- Test with sample data before production
- Use master pages for repeated elements

### CSV Preparation  
- Clean data before import (remove extra spaces)
- Use consistent naming for chairpersons
- Include all required columns even if some are empty
- Test with small datasets first

### Image Organization
- Use consistent naming conventions
- Keep images in dedicated folder
- Use high-quality images (300+ DPI)
- Maintain consistent aspect ratios

## 🔄 Version History

### v2.0 (August 30, 2025)
- **Enhanced Styling System:** Apply paragraph, table, and cell styles directly from the configuration panel
- **Style Auto-Detection:** Automatically detect and pre-select styles from template documents
- **Inline Style Creation:** Create new styles without leaving the configuration panel
- **Copy Parent Page Elements:** Preserve master page elements when duplicating pages
- **Improved UI:** Unified configuration panel with better organization
- **Bug Fixes:** Various stability improvements and error handling

### v1.0 (August 17, 2025)
- Initial Professional Release
- Complete rewrite and optimization
- Unified settings management system
- Advanced image automation
- Comprehensive error handling
- Professional reporting system
- Enterprise-grade template management

## ⚙️ Support & Updates

**Script Version:** 2.0  
**Release Date:** August 22, 2025  
**Compatibility:** InDesign CS6+
