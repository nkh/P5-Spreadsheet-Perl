package Spreadsheet::Perl::Prima::Grid;

use strict;
use warnings;
use vars qw(@ISA);
@ISA = qw(Prima::AbstractGrid);

our @cellIndents = (1,1,0,0);

sub profile_default
{
	my $def = $_[ 0]-> SUPER::profile_default;
	my %prf = (
		sheet              => undef,
		clipCells          => 1,
		cellIndents        => \@cellIndents,
		cellWidthsMap      => {},
		cellHeightsMap     => {},
		defaultCellWidth   => 80,
		defaultCellHeight  => 40,
	);
	@$def{keys %prf} = values %prf;
	return $def;
}

sub profile_check_in
{
	my ( $self, $p, $default) = @_;
	$self-> SUPER::profile_check_in( $p, $default);

	# assign default grid size
	my $sheet  = $p-> {sheet} // $default-> {sheet};
	if ( $sheet and (not exists $p-> {columns} or not exists $p-> {rows})) {
		my @range = $sheet-> GetRange;
		$p-> {columns} //= $range[0] + $cellIndents[0] + $cellIndents[2];
		$p-> {rows}    //= $range[1] + $cellIndents[1] + $cellIndents[3];
	}
}

sub init
{
	my $self = shift;
	my %profile = $self-> SUPER::init(@_);
	$self-> $_( $profile{ $_}) for qw(defaultCellHeight defaultCellWidth cellSizes sheet);
	return %profile;
}

sub on_measure
{
	my ($self, $axis, $index, $ref) = @_;

	my $map = $self-> {cellSizes};
	if ( exists $map-> [$axis]-> {$index}) {# use exists saves memory
		$$ref = $map-> [$axis]-> {$index};
	}
	else {
		$$ref = $axis ? $self-> {defaultCellHeight} : $self-> {defaultCellWidth} ;
	}
}

sub on_setextent
{
	my ($self, $axis, $index, $breadth) = @_;
	$self-> {cellSizes}-> [$axis]-> {$index} = $breadth;
}

sub on_stringify
{
	my ($self, $col, $row, $ref) = @_ ;
	$$ref = $self-> {sheet}-> Get("$col,$row") // '';
}
	
sub cellSizes
{
	return $_[0]-> {cellSizes} unless $#_;
	my ( $self, $sizes ) = @_;
	$self-> {cellSizes} = $sizes;
	$self-> reset;
	$self-> repaint;
}

sub cellIndents          { $#_ ? $_[0]-> SUPER::cellIndents(\@cellIndents) : $_[0]-> {cellIndents}     } # readonly
sub defaultCellWidth     { $#_ ? $_[0]-> {defaultCellWidth}  = $_[1]     : $_[0]-> {defaultCellWidth}  }
sub defaultCellHeight    { $#_ ? $_[0]-> {defaultCellHeight} = $_[1]     : $_[0]-> {defaultCellHeight} }
sub sheet                { $#_ ? $_[0]-> {sheet}             = $_[1]     : $_[0]-> {sheet}             }

1;
