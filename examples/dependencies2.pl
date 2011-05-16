
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;

tie my %romeo, "Spreadsheet::Perl" ;
my $romeo = tied %romeo ;
$romeo->SetName('ROMEO') ;
$romeo->{DEBUG}{INLINE_INFORMATION}++ ;
$romeo->{DEBUG}{PRINT_DEPENDENT_LIST}++ ;
#$romeo->{DEBUG}{FETCH_FROM_OTHER}++ ;
#$romeo->{DEBUG}{FETCH}++ ;

my $juliette = tie my %juliette, "Spreadsheet::Perl", NAME => 'JULIETTE' ;
$juliette->{DEBUG}{INLINE_INFORMATION}++ ;
#$juliette->{DEBUG}{PRINT_DEPENDENT_LIST}++ ;
#$juliette->{DEBUG}{FETCH_FROM_OTHER}++ ;
#$juliette->{DEBUG}{FETCH}++ ;

$romeo->AddSpreadsheet('JULIETTE', $juliette) ;
$juliette->AddSpreadsheet('ROMEO', $romeo) ;

$romeo{A1} = PerlFormula('$ss->Sum("JULIETTE!A1:A2", "A2")') ;
$romeo{A2} = 100 ;
$romeo{A3} = PerlFormula('$ss{A2}') ;
$romeo{'B1:B2'} = 10 ;
$romeo{'C1:C2'} = Formula('B1+B1') ;

$juliette{A1} = 5 ;
$juliette{A2} = PerlFormula('$ss->Sum("ROMEO!B1:B2") + $ss{"ROMEO!A2"}') ; 

print $romeo->DumpTable() ;
#print $juliette->DumpTable() ;

use Data::TreeDumper ;
my $dependencies = $romeo->GetAllDependencies('A3', 1) ;
my $title = shift @{$dependencies} ;
print DumpTree($dependencies, $title, DISPLAY_ADDRESS => 0) ;

$dependencies = $romeo->GetAllDependencies('A1', 1) ;
$title = shift @{$dependencies} ;
print DumpTree($dependencies, $title, DISPLAY_ADDRESS => 0) ;

#print DumpTree $romeo, 'romeo', MAX_DEPTH => 2 ;
#print DumpTree $juliette, 'juliette', MAX_DEPTH => 2 ;

