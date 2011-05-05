
package Spreadsheet::Perl ;

use 5.006 ;

use Carp ;
use strict ;
use warnings ;

require Exporter ;

our @ISA = qw(Exporter) ;

our %EXPORT_TAGS = 
	(
	'all' => [ qw() ]
	) ;

our @EXPORT_OK = ( @{ $EXPORT_TAGS{'all'} } ) ;

our @EXPORT ;
push @EXPORT, qw() ;

our $VERSION = '0.02' ;

#-------------------------------------------------------------------------------

sub InsertRows
{
my ($self, $start_row, $number_of_rows_to_insert) = @_ ;

confess "Invalid row '$start_row'\n" unless $start_row =~ /^\s*\d+\s*$/ ;

my %cell_list ;

for my $cell_address ($self->GetCellList())
	{
	# get all the cells for the rows under the $start_row
	my ($column, $row) = $cell_address =~ /([A-Z_]+)(\d+)/ ;

	if( $row >= $start_row)
		{
		push @{$cell_list{$row}}, $cell_address ;
		}
	}

for my $row (reverse sort keys %cell_list)
	{
	for my $cell_address (@{$cell_list{$row}})
		{
		#print "$cell_address\n" ;
		my $new_address = $self->OffsetAddress($cell_address, 0, $number_of_rows_to_insert) ; 
		
		$self->{CELLS}{$new_address} = $self->{CELLS}{$cell_address} ;
		delete $self->{CELLS}{$cell_address} ;
		}
	}
}

#-------------------------------------------------------------------------------

1 ;

__END__
=head1 NAME

Spreadsheet::Perl::QuerySet - Functions at the spreadsheet level

=head1 SYNOPSIS

  SetAutocalc
  GetAutocalc
  Recalculate
  
  SetName
  GetName
  AddSpreadsheet
  
  GetCellList
  GetLastIndexes
  GetCellsToUpdate
  
  DefineFunction
  
=head1 DESCRIPTION

Part of Spreadsheet::Perl.

=head1 AUTHOR

Khemir Nadim ibn Hamouda. <nadim@khemir.net>

  Copyright (c) 2004 Nadim Ibn Hamouda el Khemir. All rights
  reserved.  This program is free software; you can redis-
  tribute it and/or modify it under the same terms as Perl
  itself.
  
If you find any value in this module, mail me!  All hints, tips, flames and wishes
are welcome at <nadim@khemir.net>.

=cut
