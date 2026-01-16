use Test;

use CSV::Table;
use File::Temp;

use Spreadsheet::XLSX;
use Spreadsheet::XLSX::Subs;

my $host     = %*ENV<HOST>:exists ?? %*ENV<HOST> !! "unk";
my $is-local = $host eq 'bigtom' ?? True !! False;

my $debug = 0;

# Read a workbook from an existing file (can pass IO::Path or a
# Blob in the case it was uploaded).
my $file = 't/test-data/basic.xlsx';
my $workbook = Spreadsheet::XLSX.load($file);
isa-ok $workbook, Spreadsheet::XLSX;

# Get worksheets.
my $worksheets = $workbook.worksheets;
my $nsheets = $worksheets.elems;
is $nsheets, 2;
say "Workbook has $nsheets sheets" if $debug;

# Get the name of a worksheet.
my $sheetname0 = $workbook.worksheets[0].name;

# Get cell values (indexing is zero-based, done as a multi-dimensional array
# indexing operation [row ; column].
my $cells = $workbook.worksheets[0].cells;

my $aval = $cells[0;0].value;
is $aval, 'Band';
my $bval = $cells[0;1].value;
is $bval, 'Song';

if $debug {
    say .value with $cells[0;0];      # A1
    say .value with $cells[0;1];      # B1
}

# convert various data types to an XLSX object

# text attributes

# number formats

done-testing;
