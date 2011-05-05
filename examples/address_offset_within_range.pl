
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ; 

my $ss = tie my %ss, "Spreadsheet::Perl" ;

for
	(
	  ['A1:B3', 1, 1, 'A2:B3']
	)
	{
	my $offset_cell = $ss->OffsetAddress(@$_) ;
	my $offset_string = "Can't compute!" ;
	
	if(defined $offset_cell)
		{
		$offset_string = join ", ", $ss->GetCellsOffset($_->[0], $offset_cell) ;
		}
	else
		{
		$offset_cell = "Can't offset!" ;
		}
	
	print '' . (join(", ", @$_)) . " => " . $offset_cell . " offset: " . $offset_string  . "\n" ;
	}
