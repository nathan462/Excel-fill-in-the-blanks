# Excel-fill-in-the-blanks
Fill in blank cells with data from above

This addin fills in blank cells from the current selection with data from the cell above using "cel.Value = "=" & column_letter & cel.Row - 1" to provide the value of the cell above. Just highlight the range that you need filled in and click the button to fill the data.

It looks for cells that are truley empty with the IsEmpty function and cells that are effectivly empty (contain only whitespace) by using Len(Trim(cel.Value)) = 0

Installation:
The addin should be placed in C:\$user\\AppData\Roaming\Microsoft\AddIns

Open Excel and right click the ribbon to select customize ribbon

Create a New tab and Group for the addin if desired, or skip this step if not

Click the tab and group dropdowns until you find the group you want to place the addin in. Select the group and leave it highlighted

Select Macros from the dropdown on the left

Select the "Fill_in_the_blanks" macro and select add to add the adin to the group you selected.
