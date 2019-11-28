#!/usr/bin/perl

use strict;
use warnings;
use utf8;

use Excel::Writer::XLSX;
 
# Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new('write02.xlsx');
 
# Add a worksheet
my $worksheet = $workbook->add_worksheet();
 
#  Add and define a format
my $format = $workbook->add_format(); # Add a format
$format->set_bold();
$format->set_color('red');
$format->set_align('center');
 
# Write a formatted and unformatted string, row and column notation.
my $col = my $row = 0;
$worksheet->write($row, $col, 'Hi Excel!', $format);
$worksheet->write(1,    $col, '日本語もイケるのかい？');
 
# Write a number and a formula using A1 notation
$worksheet->write('A3', 1.2345);
$worksheet->write('A4', '=SIN(PI()/4)');

$workbook->close();
