
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;

$Data::TreeDumper::Displaycallerlocation =  1 ;

my $romeo = tie my %romeo, "Spreadsheet::Perl" , NAME => 'ROMEO' ;
#$romeo->{DEBUG}{FETCH}++ ;
$romeo->{DEBUG}{INLINE_INFORMATION}++ ;
$romeo->{DEBUG}{PRINT_DEPENDENT_LIST}++ ;
$romeo->{DEBUG}{DEPENDENT}++ ;
$romeo->{DEBUG}{MARK_ALL_DEPENDENT}++ ;
#$romeo->{DEBUG}{DEPENDENT_STACK_ALL}++ ;
#$romeo->{DEBUG}{FETCH_FROM_OTHER}++ ;
#$romeo->{DEBUG}{ADDRESS_LIST}++ ;

$romeo{A1} = PerlFormula('$ss->Sum("JULIETTE!A1", "A2")') ;
$romeo{A2} = 100 ;
$romeo{A3} = PerlFormula('$ss{A2}') ;

my $juliette = tie my %juliette, "Spreadsheet::Perl", NAME => 'JULIETTE' ;
#$juliette->{DEBUG}{FETCH}++ ;
$juliette->{DEBUG}{INLINE_INFORMATION}++ ;
$juliette->{DEBUG}{PRINT_DEPENDENT_LIST}++ ;
$juliette->{DEBUG}{DEPENDENT}++ ;
$juliette->{DEBUG}{MARK_ALL_DEPENDENT}++ ;
#$juliette->{DEBUG}{DEPENDENT_STACK_ALL}++ ;
#$juliette->{DEBUG}{FETCH_FROM_OTHER}++ ;

$juliette{A1} = 5 ;
$juliette{A2} = PerlFormula('$ss->Sum("ROMEO!A2", "A1")') ; 

$romeo->AddSpreadsheet('JULIETTE', $juliette) ;
$juliette->AddSpreadsheet('ROMEO', $romeo) ;

# we don't want debug output generated while dumping the 
# spreadsheet to the table
delete $romeo->{DEBUG}{DEPENDENT_STACK_ALL} ;
delete $juliette->{DEBUG}{DEPENDENT_STACK_ALL} ;

print DumpSideBySide($romeo, $juliette) ;

$romeo->Recalculate() ;
$juliette->Recalculate() ;

delete $juliette{A1} ;
print DumpSideBySide($romeo, $juliette) ;


$juliette->InsertColumns('A', 1) ;
print DumpSideBySide($romeo, $juliette) ;

$romeo->{DEBUG}{OFFSET_ADDRESS}++ ;
$juliette->{DEBUG}{OFFSET_ADDRESS}++ ;
$romeo->{DEBUG}{OFFSET_ADDRESS}++ ;
$juliette->{DEBUG}{OFFSET_ADDRESS}++ ;

$juliette->DeleteColumns('A', 1) ;
print DumpSideBySide($romeo, $juliette) ;

$juliette->DeleteColumns('A', 1) ;
print DumpSideBySide($romeo, $juliette) ;


