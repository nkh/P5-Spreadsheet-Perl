
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;

$Data::TreeDumper::Displaycallerlocation =  1 ;

my $romeo = tie my %romeo, "Spreadsheet::Perl" , NAME => 'ROMEO' ;
#$romeo->{DEBUG}{FETCH}++ ;
$romeo->{DEBUG}{INLINE_INFORMATION}++ ;
$romeo->{DEBUG}{PRINT_DEPENDENT_LIST}++ ;
$romeo->{DEBUG}{DEPENDENT}++ ;
$romeo->{DEBUG}{MARK_ALL_DEPENDENT}++ ;
#$romeo->{DEBUG}{DEPENDENT_STACK_ALL}++ ;
#$romeo->{DEBUG}{FETCH_FROM_OTHER}++ ;
#$romeo->{DEBUG}{ADDRESS_LIST}++ ;

$romeo{A1} = 100 ;
#~ $romeo{A2} = PerlFormula('$ss->{A1}') ; 
$romeo{A2} = PerlFormula(<<'EOF') ;
	my $x = $ss->{A1} ;
	use Data::Dumper ;

	print Dumper($x) ; 
	125 ;
EOF

# we don't want debug output generated while dumping the 
# spreadsheet to the table
delete $romeo->{DEBUG}{DEPENDENT_STACK_ALL} ;

print $romeo->DumpTable() ;

$romeo{A2} = PerlFormula('$ss{A1}') ; 
print $romeo->DumpTable() ;





