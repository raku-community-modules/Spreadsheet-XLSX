use Test;

use lib '.';
use Spreadsheet::XLSX;
use Spreadsheet::XLSX::Subs;

my $csv  = "t/test-data/example.csv";
my $xlsx = "simple.xlsx";

shell "./examples/csv2xlsx-simple.raku csv=$csv xlxs=$xlsx :force";

