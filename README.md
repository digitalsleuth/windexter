# Windexter

noun
win-dex-ter (ˈwɪndɛkstə˞r)

: an application with intellectual or academic interests or pretensions, geared towards the Windows Search Index

: a *poindexter* of an application designed for the Windows Search Index databases

## What Windexter Does

Windexter analyzes the contents of a `Windows.edb` (ESE database), `Windows.db` (SQLite database), or `Windows-gather.db` (SQLite database) and provides a legible output for forensic analysis in the form of an Excel spreadsheet.

The spreadsheet will contain anywhere from one to 9 tabs, based on the database provided and the options selected:

![v1-0-0](https://github.com/digitalsleuth/Windexter/raw/main/img/v1.0.0.png)

### Indexed Results

Contains the contents of the main SystemIndex_1_PropertyStore table, correlated (when available within the database) with the SystemIndex_Gthr and SystemIndex_GthrPth tables.

### Index Properties

Contains the contents of the SystemIndex_1_Properties table, when present (usually in the ESE database).

### Gather Data

Contains a correlation between the SystemIndex_Gthr and SystemIndex_GthrPth tables to display the full path for a file or folder, as well as the metadata which goes along with it.

### URL Data

Contains the filtered contents of the Indexed Results tab which shows only URL-related data.

### GPS Data

Contains the filtered contents of the Indexed Results tab which shows only GPS-related data.

### Computer Info

Contains the filtered contents of the Indexed Results tab which shows only Computer Info-related data.

### Activity

Contains the filtered contents of the Indexed Results tab which shows only Activity-related data.

### Search Summary

Contains the filtered contents of the Indexed Results tab which shows only the Search Summary-related data.

### Timeline

Contains the entirety of the Indexed Results tab, but displaying all timestamps in a single column, based on their Source (ie. System.DateCreated, System.Search.GatherTime, etc). Easily filtered and sorted.

## Project Details

This application was written in C#, utilizing the Windows Presentation Framework (WPF), and relies on [libesedb](https://github.com/libyal/libesedb/) project by Joachim Metz for easy access to the contents of the ESE database without the use of an API, and without the need to recover / repair for access. It also uses the [SQLitePCLRaw](https://www.nuget.org/packages/SQLitePCLRaw.bundle_e_sqlite3) library to enable read-only access to the SQLite databases.
