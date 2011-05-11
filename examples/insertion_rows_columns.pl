
use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;

use Carp ;
use strict ;
use warnings ;

use Data::TreeDumper ;

#-------------------------------------------------------------------------------

my $ss = tie my %ss, "Spreadsheet::Perl", NAME => 'NAME' ;

$ss->{DEBUG}{INLINE_INFORMATION}++ ;
#$ss->{DEBUG}{PRINT_ORIGINAL_FORMULA}++ ;
#$ss->{DEBUG}{FETCH}++ ;
#$ss->{DEBUG}{FETCH_VALUE}++ ;

$ss->label_column('B' => 'C-B') ;
$ss->label_row(3 => 'R-3') ;

$ss{'A1:A5'} = RangeValues(1 .. 5) ;
$ss{'B1:B5'} = PF('$ss{A1} * 2') ;
$ss{B1} = PF('$ss{B5}') ;
$ss{B3} = 'B3 is just text' ;
$ss{C1} = PF('$ss->SUM("A1:B5")') ;
$ss{C3} = PF('$ss->SUM("A3:B5")') ;
$ss{C3} = PF('$ss->SUM("A5:C5")') ;
$ss{C4} = 'some text' ;

print $ss->DumpTable() ;

$ss->InsertRows(3, 2) ;
print $ss->DumpTable() ;

$ss->InsertColumns('B', 2) ;
$ss{B1} = PF('$ss{B5} + $ss{A2}') ;
print $ss->DumpTable() ;

#print $ss->Dump() ;

