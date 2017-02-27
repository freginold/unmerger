# unmerger
unmerger is an Excel macro to separate line-delimited values in one cell into their own individual cells, without breaking up the formatting of the preceeding cells in the row.  Here is an example of what this macro does:

![screen shot]
(https://github.com/freginold/unmerger/blob/master/ss.png)

To use this function, copy the text in the .vbs file and save it as a macro in your Excel spreadsheet, then either run it from the ribbon (<b>Developer</b> > <b>Macros</b> > <b>Run</b>) or assign it a shortcut key.

### Limitations / Warnings
- As of right now, unmerger will not retain formatting for cells succeeding the target cell (to the right).
- Because it's a macro, this action can't be undone with <kbd>CTRL</kbd> <kbd>Z</kbd>.  It's a good idea to save your worksheet before trying this function.

### Tested successfully in:
- Excel 2013 / Windows 7
