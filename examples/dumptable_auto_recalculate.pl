
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;

my $romeo = tie my %romeo, "Spreadsheet::Perl", NAME => 'ROMEO' ;
$romeo->{DEBUG}{INLINE_INFORMATION}++ ;
$romeo->{DEBUG}{PRINT_DEPENDENT_LIST}++ ;
$romeo->{DEBUG}{PRINT_FORMULA_EVAL_STATUS}++ ;

#$romeo{A1} = PF('$ss{"JULIETTE!A1"} + |lm|dsq; $ss{"A2"}') ;
#$romeo{A1} = PF('$ss{"JULIETTE!A1"} + "A2"') ;
$romeo{A1} = PF('$ss{"JULIETTE!A1"} + $ss{"A2"}') ;
$romeo{A2} = 100 ;
$romeo{A3} = PF('$ss{A2}') ;

my $juliette = tie my %juliette, "Spreadsheet::Perl", NAME => 'JULIETTE' ;
$juliette->{DEBUG}{INLINE_INFORMATION}++ ;
$juliette->{DEBUG}{PRINT_DEPENDENT_LIST}++ ;
$juliette->{DEBUG}{PRINT_FORMULA_EVAL_STATUS}++ ;

$juliette{A1} = PF('$ss{"ROMEO!A2"}') ; 
$juliette{A2} = PF('$ss{"ROMEO!A1"}') ; 

$romeo->AddSpreadsheet('JULIETTE', $juliette) ;
$juliette->AddSpreadsheet('ROMEO', $romeo) ;

print $juliette->DumpTable(undef, undef, {headingText => 'Juliette'}) ;
print $romeo->DumpTable(undef, undef, {headingText => 'Romeo'}) ;
print $romeo->DumpTable(undef, undef, {headingText => 'Romeo'}) ;
#print $juliette->Dump(undef, 1) ;


