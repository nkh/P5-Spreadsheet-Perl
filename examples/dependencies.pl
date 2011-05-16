
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;
use Data::TreeDumper ;

my $ss = tie my %ss, "Spreadsheet::Perl" ;

$ss{'A1'} = PerlFormula('$ss{A2} + $ss{A5}') ;
$ss{'A2:A4'} = PerlFormula('$ss{"A3"}') ;
$ss{'A5'} = Formula('A6') ;
$ss{'A3'} = PerlFormula('$ss->Sum("A4:B5")') ;

my $dependencies = $ss->GetAllDependencies('A1', 1) ;
my $title = shift @{$dependencies} ;
print DumpTree($dependencies, $title, DISPLAY_ADDRESS => 0) ;

$ss->SetName('MY_NAME') ;
$dependencies = $ss->GetAllDependencies('A1', 1) ;
$title = shift @{$dependencies} ;
print DumpTree($dependencies, $title, DISPLAY_ADDRESS => 0) ;

delete $ss->{DEBUG}{FETCH} ;
$ss->{DEBUG}{INLINE_INFORMATION}++ ;

print $ss->DumpTable() ;

