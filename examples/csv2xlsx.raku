#!/usr/bin/env raku

use Test;
use CSV::Table;
use Spreadsheet::XLSX;

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
if not $xlsx {
    say "ERROR: The xlsx file name is missing.";
    ++$errs;
}
elsif $xlsx {
    if $xlsx.IO.r {
        if $force {
            say "NOTE: The xslx file '$xlsx' is being overwritten.";
        }
        else {
            say "ERROR: The xlsx file exists.";
            say "       Use the 'force' option to overwrite it.";
            ++$errs;
        }
     }
}
if $errs {
    say "FATAL: Too many errors.";
    exit;
}

say "Working on the input CSV file ($csv)..." if $debug;
my $ct = CSV::Table.new: :$csv;
# iterate over the rows and columns
# make sure it has a header row
unless $ct.has-header {
    say "No header row, so I lost interest...";
    say "  Exiting.";
    exit;
}
my $schar = $ct.separator;
say qq:to/HERE/;
has-header: {$ct.has-header}
field separator: '{$ct.separator}'
number of fields: {$ct.fields}
number of rows:   {$ct.rows}
number of cols:   {$ct.cols}
HERE

for 0..^$ct.fields -> $i {
    print $schar if $i;
    print $ct.field[$i];
}
say();
for 0..^$ct.rows -> $i {
    say "row $i" if $debug;
    for 0..^$ct.cols -> $j {
        print $schar if $j;
        my $s = $ct.rowcol($i, $j);
        print $s if $s.defined;
    }
    say();
}
say();

# create a new workbook and add a worksheet
my $wb = Spreadsheet::XLSX.new;
isa-ok $wb, Spreadsheet::XLSX;
my $ws = $wb.create-worksheet("data");

# fill it with data
# the header row...
for 0..^$ct.fields -> $i {
    my $text = $ct.field[$i];
    my $row-num = 0;
    my $col-num = $i;
    if $text ~~ Numeric {
        $ws.cells[$row-num;$col-num] = 
            Spreadsheet::XLSX::Cell::Number.new(value => $text);
        $ws.cells[$row-num;$col-num].style.number-format = "#,###";

        #$ws.set($row-num, $col-num, $text, :number-format("#,###"));
    }
    else {
        $ws.cells[$row-num;$col-num] = 
            Spreadsheet::XLSX::Cell::Text.new(value => $text);
        $ws.cells[$row-num;$col-num].style.bold = True;

        #$ws.set($row-num, $col-num, $text, :bold);
    }
}
# the data rows...
# note the new row numbers need adjusting if our input has a header row
#   which affects the row number by 1 (or more)
for 0..^$ct.rows -> $cti {
    my $row-num = $cti;
    if $ct.has-header {
        ++$row-num;
    }

    say "row $row-num" if $debug;
    for 0..^$ct.cols -> $j {
        #print $schar if $j;
        #print $s if $s.defined;
        my $text = $ct.rowcol($row-num, $j) // "";
        my $col-num = $j;
        if $text ~~ Numeric {
            $ws.cells[$row-num;$col-num] = 
                Spreadsheet::XLSX::Cell::Number.new(value => $text);
            $ws.cells[$row-num;$col-num].style.number-format = "#,###";

            #$ws.set($row-num, $col-num, $text, :number-format("#,###"));
        }
        else {
            $ws.cells[$row-num;$col-num] = 
                Spreadsheet::XLSX::Cell::Text.new(value => $text);
            $ws.cells[$row-num;$col-num].style.bold = True;

            #$ws.set($row-num, $col-num, $text, :bold);
        }
    }
}
=begin comment
# use convenience forms to add data
$ws.set($row-num, $col-num, $text, :bold);
$ws.set($row-num, $col-num, $number, :number-format("#,###"));
=end comment

# save it
$wb.save: $xlsx;
say "See new xlsx file: $xlsx";

