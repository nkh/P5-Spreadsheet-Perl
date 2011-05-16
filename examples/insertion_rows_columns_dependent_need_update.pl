
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;

my $romeo = tie my %romeo, "Spreadsheet::Perl", NAME => 'ROMEO' ;
$romeo->{DEBUG}{INLINE_INFORMATION}++ ;
$romeo->{DEBUG}{PRINT_DEPENDENT_LIST}++ ;

#$romeo->label_column('B' => 'C-B') ;
#$romeo->label_row(3 => 'row-3') ;

$romeo{A1} = PF('$ss{A2}') ;
$romeo{A2} = 100 ;
$romeo{A3} = PF('$ss{A2}') ;
$romeo{A4} = PF('$ss{B4}') ;
$romeo{'B1:B3'} = PF('$ss{A1}') ;

print $romeo->DumpTable() ;

$romeo->InsertRows(3, 2) ;
print $romeo->DumpTable() ;

$romeo->InsertColumns('B', 2) ;
print $romeo->DumpTable() ;

#$romeo->{DEBUG}{MARK_ALL_DEPENDENT}++ ;
#delete $romeo{A5} ;
#
#print $romeo->DumpTable() ;

$romeo->DeleteColumns('B', 1) ;
print $romeo->DumpTable() ;

$romeo->DeleteRows(2, 1) ;
print $romeo->DumpTable() ;

$romeo->DeleteColumns('A', 1) ;
print $romeo->DumpTable() ;

