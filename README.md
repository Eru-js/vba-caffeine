## vba-caffeine

Visual Basic script that prevents your system from going to sleep mode

### Enable developer mode on Excel

- in Excel click File tab and goto Options
- in Options, navigate to Customize Ribbon
- under Main tabs, ensure that the checkbox for Developer is in check

### Importing VBA module

- in Developer tab click Visual Basic to open the editor
- `Right-click` on Project-VBAProject Window and select `Import File` <br>
- navigate the `.bas` file to be imported <br>

### Integrate to Excel sheet

- in Developer tab, insert 2 button in `Sheet1`:
  - for start button
  - for stop button
- `Right-click` on the button and choose `Assign Macro`
- for start button, assign the `Move_Cursor` Macro
- for stop button, assign the `Stop_Cursor` Macro
- once you click button for start it will populate the cell in `Sheet1 column A row 1`, this is the time interval in which the macro will execute.
- to modify, click the stop button and change the time interval. Format is `hours:minutes:seconds`

### Create VBA module

- goto Visual Basic Editor <br>
- `Right-click` on Project-VBAProject Window and select `Insert` &rarr; `Module` from the context menu <br>
- start coding :heart: <br>
- to run the code: simply press `F5` to execute the whole code or `F8` to run it by line

<!-- comment -->
