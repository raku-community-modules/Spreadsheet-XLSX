[![Actions Status](https://github.com/raku-community-modules/Spreadsheet-XLSX/actions/workflows/linux.yml/badge.svg)](https://github.com/raku-community-modules/Spreadsheet-XLSX/actions) [![Actions Status](https://github.com/raku-community-modules/Spreadsheet-XLSX/actions/workflows/macos.yml/badge.svg)](https://github.com/raku-community-modules/Spreadsheet-XLSX/actions) [![Actions Status](https://github.com/raku-community-modules/Spreadsheet-XLSX/actions/workflows/windows.yml/badge.svg)](https://github.com/raku-community-modules/Spreadsheet-XLSX/actions)

NAME
====

Spreadsheet::XLSX - Work with Excel (XLSX) spreadsheets

DESCRIPTION
===========

A Raku module for working with Excel spreadsheets (XLSX format), both reading existing files, creating new files, or modifying existing files and saving the changes. Of note, it:

  * Knows how to lazily load sheet content, so if you don't look at a sheet then time won't be spent deserializing it (down to a cell level, even)

  * In the modification scenario, tries to leave as much intact as it can, meaning that it's possible to poke data into a sheet more complex than could be produced by the module from scratch

  * Only depends on the Raku `LibXML` and `Libarchive` modules (and their respective native dependencies)

This module is currently in development, and supports the subset of XLSX format features that were immediately needed for the use-case it was built for. That isn't so much, for now, but it will handle the most common needs:

  * Enumerating worksheets

  * Reading text and numbers from cells on a worksheet

  * Creating new workbooks with worksheets with text and number cells

  * Setting basic styles and number formats on cells in newly created worksheets

  * Reading a workbook, making modifications, and saving it again

  * Reading and writing column properties (such as column width)

See the `/examples directory` for some Raku scripts demonstrating its current state. Note the first example, `t/test-data/example.csv`, is a commented CSV file which can be handled easily with the published Raku package `CSV::Table` by `@tbrowder (Tom Browder)`.

SYNOPSIS
========

Reading existing workbooks
--------------------------

```raku
use Macos::NativeLib '*';
use Spreadsheet::XLSX;

# Read a workbook from an existing file (can pass IO::Path or a
# Blob in the case it was uploaded).
my $file = 'accounts.xlsx';
my $workbook = Spreadsheet::XLSX.load($file);

# Get worksheets.
my $worksheets = $workbook.worksheets;
say "Workbook has {$worksheets.elems} sheets";

# Get the name of a worksheet.
my $sheetname0 = $workbook.worksheets[0].name;

# Get cell values (indexing is zero-based, done as a multi-dimensional array
# indexing operation [row ; column].
my $cells = $workbook.worksheets[0].cells;
say .value with $cells[0;0];      # A1
say .value with $cells[0;1];      # B1
say .value with $cells[1;0];      # A2
say .value with $cells[1;1];      # B2
```

Creating new workbooks
----------------------

```raku
use Spreadsheet::XLSX;

# Create a new workbook and add some worksheets to it.
my $workbook = Spreadsheet::XLSX.new;
my $sheet-a = $workbook.create-worksheet('Ingredients');
my $sheet-b = $workbook.create-worksheet('Matching Drinks');

# Put some data into a worksheet and style it. This is how the model
# actually works (useful if you want to add styles later).
$sheet-a.cells[0;0] = Spreadsheet::XLSX::Cell::Text.new(value => 'Ingredient');
$sheet-a.cells[0;0].style.bold = True;
$sheet-a.cells[0;1] = Spreadsheet::XLSX::Cell::Text.new(value => 'Quantity');
$sheet-a.cells[0;1].style.bold = True;
$sheet-a.cells[1;0] = Spreadsheet::XLSX::Cell::Text.new(value => 'Eggs');
$sheet-a.cells[1;1] = Spreadsheet::XLSX::Cell::Number.new(value => 6);
$sheet-a.cells[1;1].style.number-format = '#,###';

# However, there is a convenience form too.
$sheet-a.set(0, 0, 'Ingredient', :bold);
$sheet-a.set(0, 1, 'Quantity', :bold);
$sheet-a.set(1, 0, 'Eggs');
$sheet-a.set(1, 1, 6, :number-format('#,###'));

# Save it to a file (string or IO::Path name).
$workbook.save("foo.xlsx");

# Or get it as a blob, e.g. for a HTTP response.
my $blob = $workbook.to-blob();
```

Class / Method reference
========================

class Spreadsheet::XLSX
-----------------------

The actual outward facing class

### has Hash $!archive

Map of files in the decompressed archive we read from, if any.

### has Spreadsheet::XLSX::ContentTypes $.content-types

The content types of the workbook.

### has Associative[Spreadsheet::XLSX::Relationships] %!relationships

Map of loaded relationships for paths. (Those never used are not in here.)

class Attribute+{<anon|2>}.new(handles => $("create-worksheet", "worksheets", "shared-strings", "styles"))
----------------------------------------------------------------------------------------------------------

The workbook itself.

### has Spreadsheet::XLSX::DocProps::Core $!core-props

Document Core Properties

### multi method load

```raku
multi method load(
    Str $file
) returns Spreadsheet::XLSX
```

Load an Excel workbook from the file path identified by the given string.

### multi method load

```raku
multi method load(
    IO::Path $file
) returns Spreadsheet::XLSX
```

Load an Excel workbook in the specified file.

### multi method load

```raku
multi method load(
    Blob $content
) returns Spreadsheet::XLSX
```

Load an Excel workbook from the specified blob. This is useful in the case it was sent over the network, and so never written to disk.

### method find-relationships

```raku
method find-relationships(
    Str $path,
    Bool :$create = Bool::False
) returns Spreadsheet::XLSX::Relationships
```

Get the relationships for a given path in the XLSX archive.

### method get-file-from-archive

```raku
method get-file-from-archive(
    Str $path
) returns Blob
```

Obtain a file from the archive. Will fail if we are not backed by an archive, or if there is no such file.

### method set-file-in-archive

```raku
method set-file-in-archive(
    Str $path,
    Blob $content
) returns Nil
```

Set the content of a file in the archive, replacing any existing content.

### method to-blob

```raku
method to-blob() returns Blob
```

Serializes the current state of the spreadsheet into XLSX format and returns a Blob containing it.

### multi method save

```raku
multi method save(
    Str $file
) returns Nil
```

Saves the Excel workbook to the file path identified by the given string.

### multi method save

```raku
multi method save(
    IO::Path $file
) returns Nil
```

Save an Excel workbook to the specified file.

### method sync-to-archive

```raku
method sync-to-archive() returns Nil
```

Synchronizes all changes to the internal representation of the archive. This is performed automatically before saving, and there is no need to explicitly perform it.

CREDITS
=======

Thanks goes to [Agrammon](https://agrammon.ch/) for making the development of this module possible.

AUTHOR
======

Jonathan Worthington

COPYRIGHT AND LICENSE
=====================

Copyright 2020 - 2024 Jonathan Worthington

Copyright 2024, 2026 Raku Community

This library is free software; you can redistribute it and/or modify it under the Artistic License 2.0.

