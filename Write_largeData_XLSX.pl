#!/usr/bin/perl

###############################################################################
#
# Example of how to use Excel::Writer::XLSX to write large data in XLSX File
# And to dynamically write the data in a new sheet if the row count exceeds 10 lakh


###############################################################################
#
# Include PERL packages.
use strict;
use Excel::Writer::XLSX;

my@records = ( [ "10001", "1/5/2011", "Jan", "Midwest", "Ami", "Binder", "94", "20" ],
                 [ "10002", "1/13/2011", "May", "West Coast", "Stevenson", "Pencil", "3", "275" ],
                 [ "10051", "2/24/2011", "Nov", "Midwest", "Jones", "Desk", "35", "4.99" ],
				 [ "10086", "3/7/2011", "Apr", "New England", "Andrews", "Pen Set", "16", "20" ],
				 [ "10450", "4/10/2011", "Dec", "Midwest", "Adams", "Ball", "20", "57.2" ],
				 [ "16001", "7/19/2011", "Mar", "West Coast", "Thompson", "Note Book", "28", "33.5" ],
				 [ "19565", "6/23/2011", "Jul", "Midwest", "Dwyer", "Scale", "15", "16" ],
				 [ "20048", "6/15/2011", "Oct", "Midwest", "Morgan", "Stickers", "96", "12" ],
				 [ "50962", "2/7/2011", "Feb", "New England", "Howard", "Clips", "52", "5.5" ],
                    );
my $outputFile 	= "D://Test.xlsx";	
# To create workbook 
my $workbook 		= Excel::Writer::XLSX->new($outputFile) or die "\nUnable to open XLSX file $outputFile in write mode: $!\n";

# To reduce memory usage 
$workbook->set_optimization();	

# To add worksheet 	
my $page = 1;
my $worksheet 	= $workbook->add_worksheet('Page_'.$page); 

# To format the cell (optional)
$worksheet->freeze_panes( 1, 0 ); 
$worksheet->set_column( 0, 7, 15 );
my $format_txt = $workbook->add_format();
my $format_amt = $workbook->add_format(num_format => '0.00');
$format_txt->set_bold();
my $date_format = $workbook->add_format( num_format => 'mm/dd/yyyy' );

# Initializing rows and columns
my $row 		= 0;
my $col 		= 0;	

# Adding header line
$worksheet->write($row, $col, 'Order ID', $format_txt);
$worksheet->write($row, $col+1, 'OrderDate', $format_txt);
$worksheet->write($row, $col+2, 'Month', $format_txt);
$worksheet->write($row, $col+3, 'Region', $format_txt);
$worksheet->write($row, $col+4, 'Employee', $format_txt);
$worksheet->write($row, $col+5, 'Item', $format_txt);
$worksheet->write($row, $col+6, 'Units', $format_txt);
$worksheet->write($row, $col+7, 'Cost', $format_txt);

my $totalCount = 0;			# To get the total number of rows written in the workbook
my $subcount = 0;			# To get the total number of rows written in the particular worsheet

# Writing rows
foreach my $record (@records){
	$row++;
	$col = 0;
	my $col = 0;
	
	# Writing columns in each row
	foreach my $data (@{$record}){
		if($col == 1){
			$worksheet->write($row,$col,$data,$date_format);
		}
		elsif($col == 7){
			$worksheet->write($row,$col,$data,$format_amt);
		}else{
			$worksheet->write($row,$col,$data);
		}			
		$col++;
	}
	$totalCount++;
	$subcount++;
	# Code to write in new sheet if the row count exceeds 10 Lakh starts (if needed)
	if($subcount == 1000000){
		$subcount = 0;
		$page++;
		# To add new worksheet 	
		$worksheet 	= $workbook->add_worksheet('Page_'.$page); 
		# To format the cell (optional)
		$worksheet->freeze_panes( 1, 0 ); 
		$worksheet->set_column( 0, 7, 15 );
		my $format_txt = $workbook->add_format();
		my $format_amt = $workbook->add_format(num_format => '0.00');
		$format_txt->set_bold();
		my $date_format = $workbook->add_format( num_format => 'mm/dd/yyyy' );		
		$row = 0;
		$col = 0;
		# Adding header line in the new sheet
		$worksheet->write($row, $col, 'Order ID', $format_txt);
		$worksheet->write($row, $col+1, 'OrderDate', $format_txt);
		$worksheet->write($row, $col+2, 'Month', $format_txt);
		$worksheet->write($row, $col+3, 'Region', $format_txt);
		$worksheet->write($row, $col+4, 'Employee', $format_txt);
		$worksheet->write($row, $col+5, 'Item', $format_txt);
		$worksheet->write($row, $col+6, 'Units', $format_txt);
		$worksheet->write($row, $col+7, 'Cost', $format_txt);
	}
	# Code to write in new sheet if the row count exceeds 10 Lakh ends
}	
	
# Closing the workbook	
$workbook->close;
	
