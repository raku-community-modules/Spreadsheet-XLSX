use Test;

use Spreadsheet::XLSX;

lives-ok {
    # convert a CSV file to an XLSX file
    my $csv-fil = "t/test-data/example.csv";
    my $xlsx = "example.xlsx";
    run "examples/csv2xlsx.raku", "csv=$csv-fil", "xlsx=$xlsx", "force", :out;

}, "example scripts run ok";

done-testing;
