
use strict ;
use warnings ;

use Devel::Size qw(total_size) ;

my $ss = {} ;
my $current_address = 'whatever' ;

# removing block below removes the error (doesn't mean it's gone but not triggered anymore)
unless(exists $ss->{CELLS}{$current_address})
	{
	$ss->{CELLS}{$current_address} = {} ;
	}

my $current_cell = $ss->{CELLS}{$current_address} ;

# not assigning to $current_cell->{SUB} removes error
($current_cell->{SUB}) = medium_sub($ss, 'whatever, even undef') ;

# this generates the error
total_size($ss) ;

#~ use Data::TreeDumper ;
#~ print DumpTree $ss;

sub medium_sub
{
my ($ss,  $formula) = @_ ;

return
	(
	sub 
		{ 
		# remove line below and you get a seg fault. Keep it and get a  double free or corruption 
		my $ss = shift ; 
		
		# uncomment line below removes the error
		my $formula = 'hi' ;
		
		local $SIG{__WARN__} = 
			sub
			{
			# not using $formula removes the error
			print "At cell formula: $formula" ;
			} ;
		}
	) ;
}
