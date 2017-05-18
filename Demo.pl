#-----------------------------------------------
#Demo.pl
#-----------------------------------------------
use strict;
use Excel::Writer::XLSX;
# Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new( 'D://Test.xlsx' );
# Add a worksheet
my $worksheet = $workbook->add_worksheet('Test_Sheet'); 
my $col = 0;
my $row = 0;
$worksheet->write( $row, $col, 'Data1',);
$row++;
$worksheet->write( $row, $col, 'Data2' );
$workbook->close;