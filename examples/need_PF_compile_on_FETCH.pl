
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;

my $romeo = tie my %romeo, "Spreadsheet::Perl",
	NAME => 'ROMEO',
	CELLS =>
		{
		A1 => { PERL_FORMULA => [undef, '$ss->Sum("JULIETTE!A1:A2", "A2")']},
		A2 => { VALUE => 100},
		A3 => { PERL_FORMULA => [undef, '$ss{A2}']},
		} ;

$romeo->{DEBUG}{INLINE_INFORMATION}++ ;
$romeo->{DEBUG}{PRINT_DEPENDENT_LIST}++ ;
$romeo->{DEBUG}{PRINT_FORMULA_EVAL_STATUS}++ ;

my $juliette = tie my %juliette, "Spreadsheet::Perl", NAME => 'JULIETTE' ;
$juliette->{DEBUG}{INLINE_INFORMATION}++ ;
$juliette->{DEBUG}{PRINT_DEPENDENT_LIST}++ ;
$juliette->{DEBUG}{PRINT_FORMULA_EVAL_STATUS}++ ;

$romeo->AddSpreadsheet('JULIETTE', $juliette) ;
$juliette->AddSpreadsheet('ROMEO', $romeo) ;


$juliette{A1} = 5 ;
$juliette{A2} = PerlFormula('$ss->Sum("ROMEO!B1:B2") + $ss{"ROMEO!A2"} + $ss{"ROMEO!A1"}') ; 
$juliette{A3} = PF('1') ;


print $romeo->DumpTable(undef, undef, {headingText => 'Romeo'}) ;
print $juliette->DumpTable(undef, undef, {headingText => 'Juliette'}) ;
#print $juliette->Dump(undef, 1) ;


