
use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;

use Carp ;
use strict ;
use warnings ;

use Data::TreeDumper ;

#-------------------------------------------------------------------------------

my $ss = tie my %ss, "Spreadsheet::Perl" ;

$ss->{DEBUG}{INLINE_INFORMATION}++ ;
#$ss->{DEBUG}{PRINT_ORIGINAL_FORMULA}++ ;
#$ss->{DEBUG}{FETCH}++ ;
#$ss->{DEBUG}{FETCH_VALUE}++ ;

$ss->label_column('B' => 'Col B') ;
$ss->label_row(5 => 'ROW 5') ;

print $ss->Dump() ;

$ss{'A1:C5'} = RangeValues(1 .. 15) ;
$ss{A5} = PF('$ss{B5}') ;
$ss{D1} = PF('$ss->SUM("A1:C5")') ;
$ss{D2} = PF('$ss->Sum("A3:D4")') ;
#$ss{C2} = PF('$ss->Sum("A3:C4") + $xyz') ;
$ss{D3} = PF('$ss{C5}') ;
$ss{D4} = PF('$ss{A1}') ;
$ss{D5} = PF('$ss{D4}') ;

print $ss->DumpTable() ;

$ss->DeleteRows(2, 2) ;
print $ss->DumpTable() ;

#print $ss->Dump() ;

