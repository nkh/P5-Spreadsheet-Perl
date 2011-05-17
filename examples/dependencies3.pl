
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;

my $romeo = tie my %romeo, "Spreadsheet::Perl" , NAME => 'ROMEO' ;
$romeo->{DEBUG}{FETCH}++ ;
$romeo->{DEBUG}{INLINE_INFORMATION}++ ;
$romeo->{DEBUG}{PRINT_DEPENDENT_LIST}++ ;
$romeo->{DEBUG}{DEPENDENT}++ ;
#$romeo->{DEBUG}{MARK_ALL_DEPENDENT}++ ;
#$romeo->{DEBUG}{DEPENDENT_STACK_ALL}++ ;
$romeo->{DEBUG}{FETCH_FROM_OTHER}++ ;

#$romeo{A1} = PerlFormula('$ss->Sum("JULIETTE!A1", "A2")') ;
$romeo{A1} = PerlFormula('$ss{"JULIETTE!A1"}') ;
$romeo{A2} = 100 ;
$romeo{A3} = PerlFormula('$ss{A2}') ;
$romeo{'B1:B2'} = 10 ;


my $juliette = tie my %juliette, "Spreadsheet::Perl", NAME => 'JULIETTE' ;
$juliette->{DEBUG}{FETCH}++ ;
$juliette->{DEBUG}{INLINE_INFORMATION}++ ;
$juliette->{DEBUG}{PRINT_DEPENDENT_LIST}++ ;
$juliette->{DEBUG}{DEPENDENT}++ ;
#$juliette->{DEBUG}{MARK_ALL_DEPENDENT}++ ;
#$juliette->{DEBUG}{DEPENDENT_STACK_ALL}++ ;
$juliette->{DEBUG}{FETCH_FROM_OTHER}++ ;

$juliette{A1} = 5 ;
$juliette{A2} = PerlFormula('$ss->Sum("ROMEO!A2", "A1")') ; 



$romeo->AddSpreadsheet('JULIETTE', $juliette) ;
$juliette->AddSpreadsheet('ROMEO', $romeo) ;



#use Data::TreeDumper ;
#my $dependencies = $juliette->GetAllDependencies('A2', 1) ;
#my $title = shift @{$dependencies} ;
#print DumpTree($dependencies, $title, DISPLAY_ADDRESS => 0) ;

print "**** Calling Recaculate()\n" ;
$romeo->Recalculate() ; #update dependents

print "**** Calling Recaculate()\n" ;
$juliette->Recalculate() ; #update dependents

# we don't want debug output generated while dumping the 
# spreadsheet to the table
delete $romeo->{DEBUG}{DEPENDENT_STACK_ALL} ;
delete $juliette->{DEBUG}{DEPENDENT_STACK_ALL} ;

#use Text::Table ;
#my $table = Text::Table->new() ;
#$table->load
#	(
#	[
#print $romeo->DumpTable(undef, undef, {headingText => 'Romeo'}) ;
#print $juliette->DumpTable(undef, undef, {headingText => 'Juliette'}) ;
#]
#	);
#print $table ;


