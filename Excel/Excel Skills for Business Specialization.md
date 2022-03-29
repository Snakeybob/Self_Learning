# What I learned from this excel specialisation

## Course 1 - Excel Skills for Business: Essentials

### Week 1

| Task                                                                              | Hotkey                 |
|-----------------------------------------------------------------------------------|------------------------|
| Toggle between Workbooks                                                          | CTRL + Tab / CTRL + F6 |
| New File/ New Workbook                                                            | CTRL+N | CMD+N         |
| Open file / Open Workbook                                                         | CTRL+O | CMD+O         |
| Close file / Close Workbook                                                       | CTRL+W | CMD+W         |
| Enter a new line of Text                                                          | ALT + Enter            |
|Takes you to the first cell that contains data in the top-left corner of the sheet | CTRL+Home              |
| Goes to the last cell in the bottom-right                                         | CTRL+End               |

#### Quick Access Toolbar
- Can be customised by adding preferred features
<br />

### Week 2 
| Task                                                                              | Hotkey                   |
|-----------------------------------------------------------------------------------|--------------------------|
| Cycle through all 4 types of cell references. (Absolute, Mixed x2, Relative)      | F4 (fn + F4 / CMD + T)   |
| Go to the previous sheet                                                          | CTRL + PgUp (CMD + PgUp) |
| Go to the next sheet                                                              | CTRL + PgDn (CMD + PgDn) |
<br />

### Week 3

#### Styles and Themes
- For consistent format

#### Format Painter
- You can double-click on the Format Painter and it will remain active as long as you need it. When you are done, simply press Escape.
<br />

### Week 4

#### Filtering
- Filter number with different rules; Filter by colour; Multiple conditions

#### Sorting
- Sort by multiple rules in order of your choice

#### Conditional Formatting
- A variety of rules and conditions to choose from

| Task                                                                                     | Hotkey                             |
|------------------------------------------------------------------------------------------|------------------------------------|
| Hide the column containing the selected cell                                             | CTRL + 0 (CMD + 0)                 |
| Hide the row containing the selected cell                                                | CTRL + 9 (CMD + 9)                 |
| Unhide the column containing the selected cell. (Select cells surrounding hidden column) | CTRL + SHIFT + 0 (CMD + SHIFT + 0) |
| Unhide the row containing the selected cell. (Select cells surrounding hidden row)       | CTRL + SHIFT + 9 (CMD + SHIFT + 9) |
| Add or remove a filter                                                                   | CTRL + SHIFT + L (CMD + SHIFT + F) |
<br />

### Week 5
#### Print Preview and to PDF
- For previewing the actual print virtually

#### Page Layout View
- For previewing and adjusting page layout

#### Orientation, Margins and Scale Setting
- To adjust the way data fitting in a page

#### Page Breaks
- Adding/adjusting page breaks manually (solid break line) and automatically (dotted break line) to your own need

#### Print Titles
- Select customised print area , Rows and columns to repeat, Page order, Gridline display

#### Headers and Footers
- To include page number, total page number, date & time, file name, worksheet name, file path, picture

| Task                                                                              | Hotkey                 |
|-----------------------------------------------------------------------------------|------------------------|
| Open Print dialogue                                                               | CTRL + P | CMD + P     |
| For selecting non adjacent cells                                                  | Shift + F8             |
<br />

### Week 6 

| Task                                                                              | Hotkey                 |
|-----------------------------------------------------------------------------------|------------------------|
| Insert new Chart sheet from selection                                             | F11 (F11*)             |

#### Change chart style & theme to be consistent or for better presentation
-----------
<br/>

## Course 2 - Excel Skills for Business: Intermediate I

### Week 1
#### 3D formulas

#### Group sheets for universal changes
- Hold Shift and click sheets
- Calculations across sheets of the same format using functions

#### Arrange all
- View all workbooks

#### Link workbooks
- Useful when all workbooks have the same format
- Changing sheet order can mess up the link formula
- Break link would clear all formula and cell references across workbooks
- Change source to repair link when a workbook location has been changed

#### Consolidate by Position
- Basically a script to run to order to do calculation across workbooks without linking workbooks

#### Consolidate by Reference
- Use to consolidate data when rows or columns are labelled but not in the same order across workooks

#### Worksheet activation
- If you have many Worksheets in your Workbook, it can be difficult to move between them to find the one you are looking for. You can use the Activate dialog to move directly to a sheet. Right-click on the arrows next to the Worksheet tabs and you will see the Activate dialog. Select the Worksheet to move to and click OK.

### Week 2
#### Automatically insert the comma between each cell reference when using a function
The CTRL (CMD) key  When you are using a function like CONCAT to join text from multiple cells, hold down the CTRL key while selecting the cells and Excel will automatically insert the comma between each cell reference for you.

#### Inserts today’s date as a fixed value
CTRL + ; (Note that this is different to the =TODAY() function because this date is fixed and will not change when you come back to your workbook tomorrow or next week.)

#### Inserts the current time as a fixed value
CTRL + SHIFT + ;| CMD + ; (Note that this is different to the =NOW() function because the time is fixed and will not change when you come back to your workbook an hour or a month later).

#### Nested Functions
We can also have a function inside another function, where a function is used as an argument, such as: =MID(A2,2,FIND(" ",A2))

#### Text to Columns
To split the text into columns. A great tool for one-off changes, and when you do not need to retain the original raw data

#### TEXTJOIN
TEXTJOIN is another function that can be used to join text together, this works well because of the following:
- You can specify once that you want a space between each word and don’t have to include a space each time like we did in CONCAT
- You now have the choice to ignore empty cells in a range.

#### Inserting a line break within text functions
- You can enter a line break inside a cell using the shortcut ALT + ENTER. However this shortcut will not work when you want to include a line break inside a text function
- We can insert a line break using the function CHAR(10), so cell C2 could be entered as either of the following functions:
=A2&CHAR(10)&B2, =CONCAT(A2,CHAR(10),B2)

#### Combining Text
- =CONCAT(), 
- =A1&B1&"C"

#### Changing Text Case
- =UPPER() - for all upper case
- =LOWER() - for all lower case
- =PROPER() - for first letter to be capital and the rest lower case

#### Extracting Text
- =LEFT() - extracts text from the left
- =MID() - extracts text from the middle, starting from certain number from the left
- =RIGHT() - extracts text from the right

#### Finding Text
=FIND()

#### Date Calculations
- =NOW() -  returns the current date and time
- =TODAY() - returns the current date
- =YEARFRAC() - calculates the fraction of the year represented by the number of whole days between two dates

### Week 3
#### Finish Editing and stay in the cell
Ctrl+Enter 

#### Paste Name
F3 - When you are writing a formula, if you cannot remember the Named Range that you are trying to use, the Paste Name dialog can help by displaying a list of all the names defined in the workbook.

#### Name Manager
Ctrl+F3 (Windows only) - You can use the Name Manager to create, modify and delete Named Ranges

#### Create Names from Selection
Ctrl+Shift+F3. Mac: Cmd+Shift+fn+F3

#### Select to the edges of the data
Ctrl+Shift+Arrows

#### See all the named ranges you have created in your table
Simply zoom out to less than 40%

#### Named Ranges
A way to give a memorable name to a single cell or range of cells. You can then use this Named Range in formulas where it functions like an absolute cell reference. It can also make your formulas more readable because the name has more meaning than a cell reference

#### Apply Names
replace cell references in formulas with defined names
- Use it with Implicit intersection operator: @

### Week 4
#### Duplicate a chart
Ctrl+D

#### Increase the font size of a chart element
Ctrl+Shift+> (This only works on Windows)

#### COUNT functions
- =COUNT() - only counts cells containing numbers
- =COUNTA() - counts the number of non-blank cells (not empty) in a cell range or the cell reference
- =COUNTBLANK() - count the number of empty cells in a given range of cells
- =COUNTIFS() - counting with Criteria

#### Adding with Criteria
=SUMIFS()

#### Sparklines
A very small line chart, typically drawn without axes or coordinates. Use sparklines to show trends in a series of values, such as:
- seasonal increases or decreases
- economic cycles
- to highlight maximum and minimum values

#### Advanced Charting
- Switch Row / Column
- Combo chart for more than one chart type in the same chart for different data series
- Secondary axis
- Change bound maximum and minimum value - affects chart scale
- Trendlines (Exponential, linear, logarithmic, Polynomial, Power, Moving Average) - forward/backward forcast
- Display R-squared value on chart

### Week 5
#### Insert Table
Ctrl+T (Cmd+T or Ctrl+T) 

#### Select table row
Shift+Space (Shift+Space)

#### Select table column
Ctrl+Space (Ctrl+Space)

#### Select table
Ctrl+A (Cmd+A)

#### Move between the corners of a table
Ctrl+.

#### Named Ranges and Tables
Named Ranges and Tables have some overlap in functionality and use, but there are some key differences.
| Named Ranges | Tables |
|--------------|--------|
| Do not automatically extend at the edges of their range| Automatically extend to include rows and columns added|
| Behaves like an absolute reference when used in a formula — when you copy the formula to another column the reference to the Named Range will remain the same| Structured References behave like a relative reference when used in a formula — when you copy the formula to another column the Structured Reference will change|
| Do not have formatting tools| Can be easily formatted|
| Named Ranges do not know about other Named Ranges — they do not work together| Provide a range of sorting and filtering tools that work with the Table as a whole unit|

#### Creating and Formatting Tables
- Table Name
- Highlight First Column, Last Column
- Total Row
- Banded Row
- Banded Column
- Filter Button
- Remove Duplicates
- Converting to Range

#### Subtotaling
Create subtotals for each selected group and create outline for different level of grouping. Only works with ranges, not tables. 

### Week 6
#### Format empty cells in a pivot table 
You can format empty cells in a pivot table by right-clicking on any cell in the table, then go to Pivot Table Options. In Layout & Format, under Format, you can choose to put a hyphen for every empty cell. A hyphen is one reasonable example of something you could add, you can add something else more appealing for yourself.

#### Pivot Tables
- Pivot table name
- Rows/Columns for text fields
- Value field
- Summarise values by
- Show values as
- Group fields
- Sort
- Filter by column or row labels

#### Show Reporting Filter Pages
Create worksheets for filtered labels

#### Inspect Data in Pivot Tables
Double click the value you want to inpect would create a snapshot worksheet with all the values that used to create that value

#### Pivoting Charts
- Switch Row / Column
- Show/hide filter buttons

#### Pivoting Slicers
Report Connections - connect to different pivot charts and pivot tables



