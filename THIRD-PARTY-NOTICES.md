# Third-party notices

Windexter and windexter-cli include or load the components below. The versions are the ones in the current build. The licences are taken from each project's published terms, so please confirm them against the packages you ship before redistributing.

| Component | Version | Used for | Licence |
|---|---|---|---|
| [libesedb](https://github.com/libyal/libesedb) | 20251130 | Reading `Windows.edb` (ESE) files. Loaded as `libesedb.dll` (Windows) or `libesedb.so` (Linux). | LGPL-3.0-or-later |
| [ClosedXML](https://github.com/ClosedXML/ClosedXML) | 0.105.1 | Writing the XLSX workbook | MIT |
| [ClosedXML.Parser](https://github.com/ClosedXML/ClosedXML.Parser) | 1.0.0 | Used by ClosedXML | MIT |
| [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) | 3.1.1 | Used by ClosedXML | MIT |
| [ExcelNumberFormat](https://github.com/andersnm/ExcelNumberFormat) | 1.1.0 | Used by ClosedXML | MIT |
| [RBush](https://github.com/viceroypenguin/RBush) | 4.0.0 | Used by ClosedXML | MIT |
| [SixLabors.Fonts](https://github.com/SixLabors/Fonts) | 1.0.0 | Used by ClosedXML to size columns | Six Labors Split License |
| [SQLitePCLRaw](https://github.com/ericsink/SQLitePCL.raw) | 3.0.2 | Reading and writing SQLite databases | Apache-2.0 |
| [SQLite](https://www.sqlite.org/) | bundled in `e_sqlite3` | The SQLite engine | Public domain |
| [CommandLineParser](https://github.com/commandlineparser/commandline) | 2.9.1 | Command line options (windexter-cli only) | MIT |
| [.NET](https://github.com/dotnet/runtime) and WPF | 10.0 | The runtime, included in self-contained builds | MIT |

## libesedb

libesedb is copyright (C) 2009-2025 Joachim Metz and is used under the GNU Lesser General Public License, version 3 or later. It is not linked into the Windexter assemblies: it is a separate library that is loaded at run time and can be replaced by another build of the same version. The source code is available from https://github.com/libyal/libesedb. If you distribute a build of libesedb that you compiled yourself, the licence requires you to make the corresponding source available as well, and to include a copy of the licence.

## Points to confirm

- **SixLabors.Fonts** is under the Six Labors Split License. It applies Apache-2.0 terms to open-source projects and to projects that only use it as a dependency of another package, and it has separate commercial terms for some closed-source uses. Windexter is open source, but check the current terms if you ever distribute it differently.
- The licence of the libesedb binary is stated here from the upstream project. The DLL's own metadata does not name a licence.
