
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ; 

my $ss = tie my %ss, "Spreadsheet::Perl" ;
$ss->SetNames("FIRST_RANGE", "A1:A2") ;

for
	(
	  ['A1', 'A1:A1', 1]  
	, ['Z9', 'Z9:Z9', 1]
	, ['ZZ1', 'ZZ1:ZZ1', 1]
	, ['AAA1', 'AAA1:AAA1', 1]
	, ['B2', 'B2:D5', 1]
	, ['D2', 'B2:D5', 1]
	, ['B5', 'B2:D5', 1]
	, ['D5', 'B2:D5', 1]
	, ['C3', 'B2:D5', 1]

	, ['A1', 'B2:D5', 0]
	, ['C1', 'B2:D5', 0]
	, ['E1', 'B2:D5', 0]
	, ['A3', 'B2:D5', 0]
	, ['E3', 'B2:D5', 0]
	, ['A6', 'B2:D5', 0]
	, ['C6', 'B2:D5', 0]
	, ['E6', 'B2:D5', 0]

	, ['A1', 'FIRST_RANGE', 1]
	, ['E6', 'FIRST_RANGE', 0]

	)
	{
	my ($cell, $range, $expected) = @{$_} ;

	printf "$cell within range $range: %d => expected $expected\n", 
		$ss->is_within_range($cell, $range) ;
	}


