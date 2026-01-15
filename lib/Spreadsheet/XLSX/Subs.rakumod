unit module Spreadsheet::XLSX::Subs;

use Spreadsheet::XLSX;
use Spreadsheet::XLSX::Cell;
use Spreadsheet::XLSX::Root;
use Spreadsheet::XLSX::XMLHelpers;
use Spreadsheet::XLSX::Types;
use LibXML::Document;
use LibXML::Element;
use LibXML::Attr;
use Spreadsheet::XLSX::Exceptions;
use Spreadsheet::XLSX::XMLHelpers;

use CSV::Table;

sub csv2xlsx(
    :$csv!,  #= the input CSV file
    :$xlsx!,
    :$force  = False,
    :$header = True, 
    :$debug,
) is export {
    =begin comment
    =end comment

    my $errs = 0;
    unless $csv.IO.r {
        say "ERROR: The csv file, $csv, does not exist or is unreadable.";
        ++$errs;
    }
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
    if $errs {
        say "FATAL: Too many errors.";
        exit;
    }

    note "Working on the input CSV file ($csv)..." if $debug;

    my $ct = CSV::Table.new: :$csv;
    # iterate over the rows and columns
    # make sure it has a header row
    unless $ct.has-header {
        say "Note this file has no header row...";
    }
    my $schar = $ct.separator;
    if $debug {
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
    } # end a debug block

    # create a new workbook and add a worksheet
    my $wb = Spreadsheet::XLSX.new;
    my $ws = $wb.create-worksheet("data");

    # fill it with data
    # the header row...
    for 0..^$ct.fields -> $i {
        my $text = $ct.field[$i];
        my $row-num = 0;
        my $col-num = $i;
        if $text ~~ Numeric {
            $ws.cells[$row-num;$col-num] = 
                number2xlsx($text, "#,###");
#               Spreadsheet::XLSX::Cell::Number.new(value => $text);
#           $ws.cells[$row-num;$col-num].style.number-format = "#,###";

            #$ws.set($row-num, $col-num, $text, :number-format("#,###"));
        }
        else {
            my @styles = ["bold"];
            $ws.cells[$row-num;$col-num] = 
                text2xlsx($text, @styles);
#               Spreadsheet::XLSX::Cell::Text.new(value => $text);
#           $ws.cells[$row-num;$col-num].style.bold = True;

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
                    number2xlsx($text, "#,###");
#                   Spreadsheet::XLSX::Cell::Number.new(value => $text);
#               $ws.cells[$row-num;$col-num].style.number-format = "#,###";

                #$ws.set($row-num, $col-num, $text, :number-format("#,###"));
            }
            else {
                my @styles = ["bold"];
                $ws.cells[$row-num;$col-num] = 
                    text2xlsx($text, @styles);
#                   Spreadsheet::XLSX::Cell::Text.new(value => $text);
#               $ws.cells[$row-num;$col-num].style.bold = True;

                #$ws.set($row-num, $col-num, $text, :bold);
            }
        }
    }
    =begin comment
    # broken:
    # use convenience forms to add data
    $ws.set($row-num, $col-num, $text, :bold);
    $ws.set($row-num, $col-num, $number, :number-format("#,###"));
    =end comment

    # save it
    my $fh = open $xlsx, :a;
    $fh.close;
    $wb.save: $xlsx;
    note "See new xlsx file: $xlsx";

} # end of sub csv2xlsx

# helpers
sub number2xlsx(
    Numeric $number,
    Str     $format is copy = "",
    --> Mu
    ) {
    # $ws.cells[$row-num;$col-num] = 
    #     Spreadsheet::XLSX::Cell::Number.new(value => $text);
    # $ws.cells[$row-num;$col-num].style.number-format = "#,###";
    my $obj = Spreadsheet::XLSX::Cell::Number.new(value => $number);

    # use any provided format
    if $format.chars {
        $obj.style.number-format = $format;
        return $obj;
    }
    # otherwise, determine from the number
    if $number ~~ /'.' (\d+) / {
        my $nd = +$0;
        my $s = "";
        $s ~= '#' for 1..$nd;

        $format = "#,###.$s";
    }
    else {
        $format = "#,###";
    }
    $obj.style.number-format = $format;
    $obj;
}
sub text2xlsx(
    Str $text,
    @styles,
#   --> Mu
    ) {
    # $ws.cells[$row-num;$col-num] = 
    #     Spreadsheet::XLSX::Cell::Text.new(value => $text);
    my $obj = Spreadsheet::XLSX::Cell::Text.new(value => $text);
    # $ws.cells[$row-num;$col-num].style.bold = True;
    for @styles -> $style {
        if $style eq 'bold' {
            $obj.style.bold = True;
        }
    }
    $obj;
}

