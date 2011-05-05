
use Spreadsheet::Perl ;

use Carp ;
use strict ;
use warnings ;

use Data::TreeDumper ;

#-------------------------------------------------------------------------------

my $ss = tie my %ss, "Spreadsheet::Perl", NAME => 'NAME' ;

#$ss->{DEBUG}{FETCH}++ ;
#$ss->{DEBUG}{FETCH_VALUE}++ ;

$ss{'A1:A5'} = RangeValues(1 .. 5) ;
$ss{'B1:B5'} = PF('$ss{A1} * 2') ;
$ss{B3} = 'B3 is just text' ;
$ss{C4} = 'some text' ;

print $ss->DumpTable() ;

$ss->InsertRows(2, 2) ;

print $ss->DumpTable() ;
#print $ss->Dump() ;

