use Test;

use CSV::Table;
use File::Temp;
use File::Find;

use Spreadsheet::XLSX;
use Spreadsheet::XLSX::Subs;

my $debug = 0;

my $host     = %*ENV<HOST>:exists ?? %*ENV<HOST> !! "unk";
my $is-local = $host eq 'bigtom' ?? True !! False;
# do we need a tmpdir? yes, indeed
my $tmpdir = tempdir;
my $tdir = mkdir "$tmpdir/dir";
$tdir.IO.chmod(0o777);

# convert a CSV file to an XLSX file
my $edir = "t/test-data";
my @csv-fils = find :dir($edir), :type<file>, :name(/:i '.' csv $/);
for @csv-fils -> $csv-fil {
    # skip the headerless file
    next if $csv-fil ~~ /headerless/;
    say "Found csv file: $csv-fil" if 1 or $debug;
    my $csv = $csv-fil.basename;
    say "  its basename: $csv" if 1 or $debug;
    my $xlsx-base = $csv;
    $xlsx-base ~~ s/:i '.' csv $/.xlsx/;

    my $xlsx;
    if $is-local {
        $xlsx = $xlsx-base;
    }
    else {
        $xlsx = "$tdir/$xlsx-base";
    }

    =begin comment
    # do i need this? yes, for later tests to compare cells
    my $ct  = CSV::Table.new: :$csv;
    isa-ok $ct, CSV::Table;
    my $ct  = CSV::Table.new: :$csv;
    isa-ok $ct, CSV::Table;
    =end comment

    # The plan is, for each csv file, convert it to an xlsx file
    # and check valid output.
    lives-ok {
        my $wb = csv2xlsx :csv($csv-fil);
        isa-ok $wb, Spreadsheet::XLSX;
        $wb.save: $xlsx;
        my $res = shell "file $xlsx", :out;
        my $typ = $res.out.slurp(:close);
        $typ .= trim;
        say "File type is: $typ" if 1 or $debug;
        cmp-ok $typ, '~~', "$xlsx: Microsoft Excel 2007+",
            "got correct file type: Excel";
    }, "test output file '$xlsx";

} # end @csv loop

done-testing;
=finish 

k
my $csv = "t/test-data/example.csv";
my $ct  = CSV::Table.new: :$csv;
isa-ok $ct, CSV::Table;

my $xlsx-base;
$xlsx-base = "example0.xlsx";
my $xlsx;

if $is-local {
    $xlsx = $xlsx-base;
}
else {
    $xlsx   = "$tdir/$xlsx-base";
}

lives-ok {
    my $wb = csv2xlsx :$csv;
    isa-ok $wb, Spreadsheet::XLSX;
    $wb.save: $xlsx;
    my $res = shell "file $xlsx", :out;
    my $typ = $res.out.slurp(:close);
    $typ .= trim;
    say "File type is: $typ" if 1 or $debug;
    cmp-ok $typ, '~~', "$xlsx: Microsoft Excel 2007+",
        "got correct file type: Excel";
}, "test output file '$xlsx-base";

done-testing;
=finish

# CSV contents:
=begin comment
last, first, age, country  # header row
Lee, Sara, 35, US
Jones, Billy, 42, UK
Wabash, Winston,, CA
Powell, Bob          # no more data
=end comment

$xlsx-base = "example.xlsx";

# csv2xlsx-hardway.raku
# csv2xlsx-simple.raku

lives-ok {

    # convert a CSV file to an XLSX file
    # note the example.csv file has comments which can
    # be handled with the help of Raku package CSV::Table
    my $csv-fil = "t/test-data/example.csv";

    my $xlsx;
    if $is-local {
        $xlsx = "$xlsx-base";
    }
    else {
        $xlsx   = "$tdir/$xlsx-base";
    }

    my $proc;
    my $eprog = "examples/csv2xlsx-hardway.raku";

    #if not $is-local {
    if 1 {
        # with run we need a special trick to write to a file, per Google AI:
        my $fh = open :w, $xlsx.IO or die "Unable to open path '$xlsx'";
        $proc = run($eprog, "csv=$csv-fil",
                   "xlsx=$xlsx", "debug", "force", :out($fh), :err);
    }
    else {
        $proc = run($eprog, "csv=$csv-fil",
                   "xlsx=$xlsx", "debug", "force", :out, :err);
    }

    my $outstr = $proc.out.slurp(:close);
    my $errstr = $proc.err.slurp(:close);
    my $e   = $proc.exitcode;
    is $e, 1, "good exitcode $e";

    if not $is-local {
        my $res = shell "file $xlsx", :out;
        my $typ = $res.out.slurp(:close);
        $typ .= trim;
        say "File type is: $typ" if $debug;
        cmp-ok $typ, '~~', "$xlsx: Microsoft Excel 2007+",
        "got correct file type: Excel";
    }

}, "example scripts run ok; test output file '$xlsx-base";

lives-ok {

    # convert a CSV file to an XLSX file
    # note the example.csv file has comments which can
    # be handled with the help of Raku package CSV::Table
    my $csv-fil = "t/test-data/example.csv";

    my $xlsx;
    if $is-local {
        $xlsx = "$xlsx-base";
    }
    else {
        $xlsx   = "$tdir/$xlsx-base";
    }

    my $proc;
    my $eprog = "examples/csv2xlsx-easyway.raku";

    if not $is-local {
        # with run we need a special trick to write to a file, per Google AI:
        my $fh = open :w, $xlsx.IO or die "Unable to open path '$xlsx'";
        $proc = run($eprog, "csv=$csv-fil",
                   "xlsx=$xlsx", "debug", "force", :out($fh), :err);
    }
    else {
        $proc = run($eprog, "csv=$csv-fil",
                   "xlsx=$xlsx", "debug", "force", :out, :err);
    }

    my $outstr = $proc.out.slurp(:close);
    my $errstr = $proc.err.slurp(:close);
    my $e   = $proc.exitcode;
    is $e, 1, "good exitcode $e";

    if not $is-local {
        my $res = shell "file $xlsx", :out;
        my $typ = $res.out.slurp(:close);
        $typ .= trim;
        say "File type is: $typ" if $debug;
        cmp-ok $typ, '~~', "$xlsx: Microsoft Excel 2007+",
        "got correct file type: Excel";
    }

}, "example scripts run ok; test output file '$xlsx-base";

done-testing;
=finish

=begin comment
lives-ok {
    # convert a CSV file to an XLSX file

    my $csv-fil = "t/test-data/example-pipes.csv";
    my $xlsx = "example-pipes.xlsx";

    run "examples/csv2xlsx.raku", "csv=$csv-fil", "xlsx=$xlsx", "force", :out;

    my $res = shell "file $xlsx", :out;
    my $typ = $res.out.slurp(:close);
    $typ .= trim;
    say "File type is: $typ" if $debug;
    cmp-ok $typ, '~~', "$xlsx: Microsoft Excel 2007+",
        "got correct file type: Excel";
}, "example pipes scripts run ok";
=end comment

=begin comment
lives-ok {
    # convert a CSV file to an XLSX file

    my $csv-fil = "t/test-data/example-semis.csv";
    my $xlsx = "example-semis.xlsx";

    run "examples/csv2xlsx.raku", "csv=$csv-fil", "xlsx=$xlsx", "force", :out;

    my $res = shell "file $xlsx", :out;
    my $typ = $res.out.slurp(:close);
    $typ .= trim;
    say "File type is: $typ" if $debug;
    cmp-ok $typ, '~~', "$xlsx: Microsoft Excel 2007+",
        "got correct file type: Excel";
}, "example semis scripts run ok";
=end comment

done-testing;
