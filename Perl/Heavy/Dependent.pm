
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
push @EXPORT, qw( ) ;

our $VERSION = '0.13' ;

#-------------------------------------------------------------------------------

sub FindDependent
{
my ($self, $current_cell, $start_cell) = @_ ;

if(exists $self->{DEPENDENT_STACK} && @{$self->{DEPENDENT_STACK}})
	{
	my $last_dependent = @{$self->{DEPENDENT_STACK}}[-1] ;
	
	# keep our own copy to keep our sanity whenmultiple cells>spreadsheets point 
	# at the same thing!
	my $dependent_spreadsheet = $last_dependent->[0] ;
	my $dependent_cell = $last_dependent->[1] ;
	
	my $dependent_name = $dependent_spreadsheet->GetName() . "!$dependent_cell" ;
	
	if($self->{DEBUG}{DEPENDENT})
		{
		$current_cell->{DEPENDENT}{$dependent_name}{DEPENDENT_DATA} = 	[$dependent_spreadsheet, $dependent_cell] ;

		$current_cell->{DEPENDENT}{$dependent_name}{COUNT}++ ;
		}
	else
		{
		$current_cell->{DEPENDENT}{$dependent_name}{DEPENDENT_DATA} = [$dependent_spreadsheet, $dependent_cell] ;
		}
	}
}

#-------------------------------------------------------------------------------

sub MarkDependentForUpdate
{
my ($self, $current_cell, $cell_name, $level) = @_ ;

$level ||= 1 ;

return unless exists $current_cell->{DEPENDENT} ;

push @{$self->{DEPENDENT_STACK}}, [$self, $cell_name] ;

if(exists $current_cell->{CYCLIC_FLAG})
	{
	my $dh = $self->{DEBUG}{ERROR_HANDLE} ;

	my $full_cell_name = $self->GetName() . '!' . $cell_name ; 
	print $dh $self->DumpDependentStack("Cyclic dependency at '$full_cell_name' while marking cells for update\n") ;

	return ;
	}

$current_cell->{CYCLIC_FLAG}++ ;

if($level == 1 && (exists $self->{DEBUG}{MARK_ALL_DEPENDENT} || exists $self->{DEBUG}{MARK_DEPENDENT}{$cell_name}))
  {
  my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
  my $full_cell_name = $self->GetName() . '!' . $cell_name ; 
  print $dh "Marking dependents for update at cell: $full_cell_name\n" ;
  }

for my $dependent_name (keys %{$current_cell->{DEPENDENT}})
	{
	my $dependent = $current_cell->{DEPENDENT}{$dependent_name}{DEPENDENT_DATA} ;
	my ($dependent_spreadsheet, $dependent_cell_name) = @$dependent ;
	
	if(exists $dependent_spreadsheet->{CELLS}{$dependent_cell_name})
		{
		$dependent_spreadsheet->{CELLS}{$dependent_cell_name}{NEED_UPDATE}++ ;

		if(exists $self->{DEBUG}{MARK_ALL_DEPENDENT} || exists $self->{DEBUG}{MARK_DEPENDENT}{$cell_name})
			{
			my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
			my $full_dependent_cell_name = $dependent_spreadsheet->GetName() . '!' . $dependent_cell_name ; 
			print $dh ('   ' x $level) . "$full_dependent_cell_name needs update\n" ;
			}

		$dependent_spreadsheet->MarkDependentForUpdate($dependent_spreadsheet->{CELLS}{$dependent_cell_name}, $dependent_cell_name, $level+1) ;
		}
	else
		{
		my $full_cell_name = $self->GetName() . '!' . $cell_name ; 
		my $full_dependent_cell_name = $dependent_spreadsheet->GetName() . '!' . $dependent_cell_name ; 

		die "Marking dependents for update at $full_cell_name depend $full_dependent_cell_name does't exist"  ;
		}
	}

pop @{$self->{DEPENDENT_STACK}} ;

delete $current_cell->{CYCLIC_FLAG} ;
}

#-------------------------------------------------------------------------------

sub InvalidateCellInDependent
{

my ($self, $dependent_name) = @_ ;

#TODO: also invalidate in all known spreadsheets

for my $current_address ($self->GetCellList())
	{
	if(exists $self->{CELLS}{$current_address}{DEPENDENT})
		{
		delete $self->{CELLS}{$current_address}{DEPENDENT}{$dependent_name} ;
		}
	}
}


#-------------------------------------------------------------------------------
1 ;

__END__
=head1 NAME

Spreadsheet::Perl::Heavy::Dependent - manipulation of cell dependents

=head1 DESCRIPTION

Part of Spreadsheet::Perl.

=head1 AUTHOR

Khemir Nadim ibn Hamouda. <nadim@khemir.net>

  Copyright (c) 2004, 2011 Nadim Ibn Hamouda el Khemir. All rights
  reserved.  This program is free software; you can redis-
  tribute it and/or modify it under the same terms as Perl
  itself.
  
If you find any value in this module, mail me!  All hints, tips, flames and wishes
are welcome at <nadim@khemir.net>.

=cut

