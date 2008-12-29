
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Data::TreeDumper ;
use Devel::Size qw(size total_size) ;

# error 1
#~ my $ss = tie my %ss, "Spreadsheet::Perl" ;
#~ $ss{A1} = FetchFunction('', sub{$ss->Dump() ;}) ;
#~ my $x = $ss{A1} ;

#~ # error 2
my $ss = tie my %ss, "Spreadsheet::Perl" ;
$ss{A1} = PerlFormula('$ss{A5} * 5') ;
print $ss->Dump() ;

