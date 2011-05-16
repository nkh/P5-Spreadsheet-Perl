
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;

my $romeo = tie my %romeo, "Spreadsheet::Perl", NAME => 'ROMEO' ;
$romeo->{DEBUG}{INLINE_INFORMATION}++ ;
$romeo->{DEBUG}{PRINT_DEPENDENT_LIST}++ ;

$romeo{A1} = PF('$ss{A2}') ;
$romeo{A2} = 100 ;
$romeo{A3} = PF('$ss{A2}') ;
$romeo{'B1:B3'} = PF('$ss{A1}') ;

print $romeo->DumpTable() ;

delete $romeo{A3} ;
print $romeo->DumpTable() ;

delete $romeo{A2} ;
print $romeo->DumpTable() ;

delete $romeo{B2} ;
print $romeo->DumpTable() ;
