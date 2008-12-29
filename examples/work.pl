
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;


my $ss = tie my %ss, "Spreadsheet::Perl" ;
$ss{'A1:A2'} = 5 ;
$ss{A3} = PF('$ss->Sum("A1:A2")') ;


#$ss->{DEBUG}{FETCH}++;
$ss->{DEBUG}{FETCH} = sub{print "At '$_[2]':\n" . $_[0]->DumpDependentStack()} ;

print $ss->DumpTable() ;

	
