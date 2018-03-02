# xl DuplicateManager
[Excel-Add-In (.xlam)](https://marco-krapf.de/xl-duplicate-manager/) - Tool for finding, deleting and exporting duplicates in an Excel worksheet with lots of extra features

## Version history

### Version 2.0 (02 March 2018)
* Additional search mode: merging of non-contiguous columns
* Additional deletion mode: deleting complete lines containing duplicates
* Highlighting of unique values (values that occur only once) on the worksheet
* Automatic highlighting of duplicates, multiple values or unique values possible
* Duplicates can either be highlighted in solid red or with separate colours for the same values
* Output of unique values, optionally as whole lines
* Status area redesigned: permanent display of the number of duplicates, unique values, different values and deleted duplicates
* Screen refresh can be turned off for complex calculations
* Progress bar at the hourglass while calculating
* Tool performance significantly improved
* Tool settings can be saved automatically
* Smaller GUI with functions grouped in tabs
* GUI can be switched between English and German
* Function removed: scrolling on the worksheet
* Function removed: generation of demo data
* Bugfixes

### Version 1.2 (05 May 2016)
* Selection of entire columns on the worksheet possible
* Activation of pop-ups with hints possible
* GUI usability improved
* Bugfixes

### Version 1.1 (24 April 2016)
* Generation of demo data to test the capabilities of the tool possible (removed in version 2.0)
* Bugfixes

### Version 1.0 (05 April 2016)
* Selection of multiple ranges on the worksheet while holding down the Ctrl key
* Ignore upper/lower-case letters and/or blanks for the search if needed
* Coloured highlighting of duplicates on the worksheet
* Ausgeben der Duplikate auf einem neuen Tabellenblatt, auch als ganze Zeile
* Output the duplicates or optionally the whole lines on a new worksheet
* Delete and restore the duplicates on the worksheet, optionally with compression
* Springen zu einem einzelnen Duplikat durch Anklicken im Duplikatfenster
* Jump to a single duplicate by clicking in the duplicate window
* Non-destructive behavior as long as the selection is not changed and the tool is not closed
