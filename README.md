# unmerger
unmerger is an Excel macro to separate line-delimited values in one cell into their own individual cells, without breaking up the formatting of the preceeding cells in the row.  Here is an example of what this macro does:

![screen shot](https://github.com/freginold/unmerger/blob/master/ss.png)

### Usage:
To use this function, copy the text in the .vbs file and save it as a macro in your Excel spreadsheet, then either run it from the ribbon (<b>Developer</b> > <b>Macros</b> > <b>Run</b>) or assign it a shortcut key.

Cells to the left of the affected cell will retain their row height and formatting.  Cells to the right will also retain their row height and formatting, until three empty cells in a row are passed; then all subsequent cells will take on the row height of the affected cell. That value can be adjusted by changing the line `Do Until empties > 2` to replace `2` with the maximum number of adjacent empty cells, minus 1.

### Limitations / Warnings:
- Because it's a macro, this action can't be undone with <kbd>CTRL</kbd> <kbd>Z</kbd>.  It's a good idea to save your worksheet before trying this function.
- Using this macro on multiple columns in the same row will probably not work as expected. It's best to not use it more than once on the same row.

### Tested successfully in:
- Excel 2007 / Windows 7
- Excel 2013 / Windows 7
