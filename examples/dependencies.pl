
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;
use Data::TreeDumper ;

my $ss = tie my %ss, "Spreadsheet::Perl", NAME=> 'TEST' ;

#$ss->{DEBUG}{FETCH}++ ;

$ss{'A1'} = PerlFormula('$ss{A2} + $ss{A5}') ;
$ss{'A2:A5'} = PerlFormula('$ss{"A3"}') ;
$ss{'A3'} = PerlFormula('$ss->Sum("A4:B5")') ;

my $dependencies = $ss->GetAllDependencies('A1', 1) ;
my $title = shift @{$dependencies} ;
print DumpTree($dependencies, $title, DISPLAY_ADDRESS => 0) ;

delete $ss->{DEBUG}{FETCH} ;
$ss->{DEBUG}{INLINE_INFORMATION}++ ;

print $ss->DumpTable() ;

__DATA__

├─ 0 = TEST!A1  [S1]
├─ 1 = TEST!A2  [S2]
├─ 2 =       TEST!A3  [S3]
├─ 3 =          TEST!A4  [S4]
├─ 4 =             TEST!A5  [S5]
├─ 5 =                TEST!A6  [S6]
├─ 6 =          TEST!A5  [S7]
├─ 7 =          TEST!B4  [S8]
└─ 8 =          TEST!B5  [S9]


