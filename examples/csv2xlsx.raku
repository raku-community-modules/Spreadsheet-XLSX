#!/usr/bin/env raku

if not @*ARGS {
    print qq:to/HERE/;
    Usage: {$*PROGRAM.basename} <args> [options]

    Args (all are required):
      csv=X   - Where X is the CSV file to be used as input
      xlsx=X  - Where X is the XLSX file to be created

    Options:
      force  - overwrite an existing file
      debug	
    HERE
    exit;
}

my ($csv, $xlsx, $debug, $force) = "", "", False, False;

for @*ARGS {
    when /^:i 'csv='(\S+) / {
        $csv = ~$0;
    }
    when /^:i 'xlsx='(\S+) / {
        $xlsx = ~$0;
    }
    when /^:i d / {
        $debug = True;
    }
    when /^:i f / {
        $force = True;
    }
    default {
        die qq:to/HERE/;
        FATAL: Unknown arg '$_'
        HERE
    }
}

my $errs = 0;
unless $csv.IO.r {
    say "ERROR: The csv file, $csv, does not exist or is unreadable.";
    ++$errs;
}
if not $xlsx.defined {
    say "ERROR: The xlsx file name is missing.";
    ++$errs;
}
elsif $xlsx.IO.e {
    if $force {
        say "NOTE: The xslx file '$xlsx' is being overwritten.";
    }
    else {
        say "ERROR: The xlsx file exists.";
        say "       Use the 'force' option to overwrite it.";
        ++$errs;
    }
}
