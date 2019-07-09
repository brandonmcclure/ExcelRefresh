# ExcelRefresh.exe (XLRefresh in C#)
Command line utility to refresh Excel documents and their external connections, query tables and pivot tables.

Example usage.  Use Window's Task Scheduler to run a .bat which updates Excel files and then copies them to a public location.  

As of Excel 2013, the query table method seems to have been replaced with an external connection method.  Leaving the query table just in case.

  -f, --file            Required. Input file to be processed.                          
                                                                                       
  -m, --Macros          The worksheet macros to run. Example: -m sheet1.someMacro (sheet2.otherMacro)                           
                                                                                       
  -d, --verbose         (Default: False) Prints all messages to standard output.                                                        
                                                                                       
  -v, --visible         (Default: False) Shows Excel while update is running.                                   
                                                                                       
  -p, --pivot-tables    (Default: False) Refresh Pivot-tables.                         
                                                                                       
  -q, --query-tables    (Default: False) Refresh query-tables. (Pre Excel 2013)        
                                                                                       
  -c, --connections     (Default: False) Refresh External connections. (Excel 2013)                                                          


## Contribution

Forked from [alapolloni](https://github.com/alapolloni/ExcelRefresh), which was based on Perl program/library originally written by [CTBROWN](http://cpansearch.perl.org/src/CTBROWN/Win32-Excel-Refresh-0.02/extras/XLRefresh.pl) 
 
## License
See [LICENSE](/LINCENSE.md)