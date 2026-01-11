#!/usr/bin/env raku

if not @*ARGS {
    print qq:to/HERE/;
    Usage: {$*PROGRAM.basename} <args> [options]

    Args
      :\$csv  - the CSV file to be used as input
      :\$xlsx - the XLSX file to be created

    Options:
      force  - overwrite an existing file
      debug	
    HERE
    exit;
}

