# Excel_Writer_XLSX
    To handle large amount of data in XLSX file using Excel-Writer-XLSX perl module.
    And to dynamically write the data in a new sheet if the row count exceeds 10 lakh

# Prerequisite
        Perl Module
        Excel-Writer-XLSX module
    
# References
	http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX/lib/Excel/Writer/XLSX.pm
    
# Installation of Excel-Writer-XLSX module
    1. Download Excel-Writer-XLSX-0.95.tar.gz from http://search.cpan.org/~jmcnamara/Excel-Writer-XLSX/lib/Excel/
       Writer/XLSX.pm
        
    2. Place the Unzipped & untaredperl Excel module in lib folder of the Perl installed directory 
       (i.e., \\<perl_Installed_ directory>\Perl\lib). After placing the file the folder structure will be as 
       \\<perl_Installed_directory>\Perl\lib\Excel\Writer\..  
    
    Run demo.pl Script to check whether the Excel Writer is installed properly or not.
    
        If the demo.pl runs without any error and if the XLSX file is created as expected then 
        the EXCEL Writer module is installed properly.
        
        But if demo.pl returns any error like “Can't locate object method "newdir" via package "File::Temp" 
        at <perl_Installed_directory>/Perl/lib/Excel/Writer/XLSX/Workbook.pm line 1033”. Then the Workbook.pm 
        module installed is not   supported. In that case follow the below steps:
            1. Open “Workbook.pm” from the <perl_Installed_directory>/Perl/lib/Excel/Writer/XLSX path.
            2. Comment out the below line
                my $tempdir  = File::Temp->newdir( DIR => $self->{_tempdir} ); (line no: 1033)
            3. Add the below line after the commented line
                my $tempdir = File::Temp::tempdir(DIR => $self->{_tempdir});  
       The above two steps will solve the newdir error in the Workbook perl module.

# Handling large amount of data in XLSX
    
    Write_largeData_XLSX.pl	 -----> To use Excel::Writer::XLSX to write large data in XLSX File

    To write large amount of data in XLSX and to handle large data and to reduce memory usage set_optimization() 
    method is used.
             $workbook->set_optimization();
	     
    This optimization turned on a row of data is written and then discarded when a cell in a new row is added via 
    one of the Worksheet write_*() methods. As such data should be written in sequential row order once the 
    optimization is turned on.
    This method must be called before any calls to add_worksheet().


