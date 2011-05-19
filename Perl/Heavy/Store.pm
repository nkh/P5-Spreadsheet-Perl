
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

#-------------------------------------------------------------------------------

sub STORE 
{
my $self    = shift ;
my $address = shift ;
my $value   = shift ;

# when inserting and deleting, dependents are update so 
# the minimum amout of recalculations are done
# still this sub is called to set different fields in the cell
my $do_not_mark_dependents_for_update = shift ; 

# inter spreadsheets references
my $original_address = $address ;
my $ss_reference ;

my ($cell_or_range, $is_cell, $start_cell, $end_cell) = $self->CanonizeAddress($address) ;

($ss_reference, $address) = $self->GetSpreadsheetReference($cell_or_range) ;

if(defined $ss_reference)
	{
	if($ss_reference == $self)
		{
		#~ print "fine, it's us" ;
		}
	else
		{
		if($self->{DEBUG}{REDIRECTION})
			{
			my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
			print $dh $self->GetName() . " Store redirected to spreadsheet '$original_address'.\n" ;
			}
			
		return($ss_reference->Set($address, $value)) ;
		}
	}
else
	{
	confess "Can't find Spreadsheet object for address '$address'.\n." ;
	}
	
if($self->{DEBUG}{STORE})
	{
	my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
	#~ print $dh "Storing To '$address' @ @{[join ':', caller()]}\n" ;
	print $dh "Storing To '$address'\n" ;
	}
	
# Set the value in the current spreadsheet
for my $current_address ($self->GetAddressList($address))
	{
	unless(exists $self->{CELLS}{$current_address})
		{
		$self->{CELLS}{$current_address} = {} ;
		}
	
	my $current_cell = $self->{CELLS}{$current_address} ;
	
	if($self->{DEBUG}{STORED})
		{
		$current_cell->{STORED}++ ;
		}
		
	# triggers
	if(exists $self->{DEBUG}{STORE_TRIGGER}{$current_address})
		{
		if('CODE' eq ref $self->{DEBUG}{STORE_TRIGGER}{$current_address})
			{
			$self->{DEBUG}{STORE_TRIGGER}{$current_address}->($self, $current_address, $value) ;
			}
		else
			{
			if(exists $self->{DEBUG}{STORE_TRIGGER_HANDLER})
				{
				$self->{DEBUG}{STORE_TRIGGER_HANDLER}->($self, $current_address, $value) ;
				}
			else
				{
				my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
				my $value_text = "$value" if defined $value ;
				$value_text    = 'undef' unless defined $value ;
				print $dh "Storing cell '$current_address' => $value_text\n" ;
				}
			}
		}
		
	# validators
	my $value_is_valid = 1 ;
	
	unless(defined $value && ref $value =~ /^Spreadsheet::Perl/)
		{
		my $cell_validators = $current_cell->{VALIDATORS} if(exists $current_cell->{VALIDATORS}) ;
		
		for my $validator_data (@{$self->{VALIDATORS}}, @$cell_validators)
			{
			if(0 == $validator_data->[1]($self, $current_address, $current_cell, $value))
				{
				$value_is_valid = 0 ;
				last ;
				}
			}
		}
		
	if($value_is_valid)
		{
		$self->MarkDependentForUpdate($current_cell, $address) unless $do_not_mark_dependents_for_update ;

		$self->InvalidateCellInDependent($self->GetName() . '!' . $current_address) ;

		$current_cell->{DEFINED_AT} = [caller] if(exists $self->{DEBUG}{DEFINED_AT}) ;
		
		for (ref $value)
			{
			/^Spreadsheet::Perl::Cache$/ && do
				{
				$current_cell->{CACHE} = $$value ;
				last ;
				} ;
				
			(
			   /^Spreadsheet::Perl::Formula$/
			|| /^Spreadsheet::Perl::PerlFormula$/
			) && do
				{
				delete $current_cell->{VALUE} ;
				
				my $sub_generator = $value->[0] ;
				my $formula = $value->[1] ;
				
				if(/^Spreadsheet::Perl::Formula$/)
					{
					$current_cell->{FORMULA} = $value ; # should we compile and check the formula directly?
					delete $current_cell->{PERL_FORMULA} ;
					}
				else
					{
					$current_cell->{PERL_FORMULA} = $value ;
					delete $current_cell->{FORMULA} ;
					}
				
				$current_cell->{FETCH_SUB_ARGS} = [(@$value)[2 .. (@$value - 1)]] ;
				$current_cell->{NEED_UPDATE}    = 1 ;
				$current_cell->{ANCHOR}         = $address ;
				($current_cell->{FETCH_SUB}, $current_cell->{GENERATED_FORMULA}) = $sub_generator->(
															  $self
															, $current_address
															, $address #anchor
															, $formula
															) ;
				last ;
				} ;
				
			/^Spreadsheet::Perl::Format$/ && do
				{
				@{$current_cell->{FORMAT}}{keys %$value} = values %$value ;
				last ;
				} ;
				
			/^Spreadsheet::Perl::Validator::Add$/ && do
				{
				push @{$current_cell->{VALIDATORS}}, [$value->[0], $value->[1]] ;
				last ;
				} ;
			
			/^Spreadsheet::Perl::Validator::Set$/ && do
				{
				$current_cell->{VALIDATORS} = [[$value->[0], $value->[1]]] ;
				last ;
				} ;
			
			/^Spreadsheet::Perl::StoreFunction$/ && do
				{
				delete $current_cell->{VALUE} ;
				
				$current_cell->{STORE_SUB_INFO} = $value->[0] ;
				$current_cell->{STORE_SUB}      = $value->[1] ;
				$current_cell->{STORE_SUB_ARGS} = [ @$value[2 .. (@$value - 1)] ] ;
				last ;
				} ;
				
			/^Spreadsheet::Perl::FetchFunction$/ && do
				{
				delete $current_cell->{VALUE} ;
				
				$current_cell->{FETCH_SUB_INFO}    = $value->[0] ;
				$current_cell->{FETCH_SUB}         = $value->[1] ;
				$current_cell->{FETCH_SUB_ARGS}    = [ @$value[2 .. (@$value - 1)] ] ;
				$current_cell->{NEED_UPDATE} = 1 ;
				last ;
				} ;
				
			/^Spreadsheet::Perl::UserData$/ && do
				{
				$current_cell->{USER_DATA} = {@$value} ;
				last
				} ;
				
			/^Spreadsheet::Perl::StoreOnFetch$/ && do
				{
				$current_cell->{STORE_ON_FETCH}++ ;
				last
				} ;
				
			/^Spreadsheet::Perl::DeleteFunction/ && do
				{
				$current_cell->{DELETE_SUB_INFO}    = $value->[0] ;
				$current_cell->{DELETE_SUB}         = $value->[1] ;
				$current_cell->{DELETE_SUB_ARGS}    = [ @$value[2 .. (@$value - 1)] ] ;
				last
				} ;
				
			/^Spreadsheet::Perl::Reference$/ && do
				{
				$current_cell->{IS_REFERENCE}  = 1 ;
				$current_cell->{REF_SUB_INFO}  = $value->[0] ;
				$current_cell->{REF_STORE_SUB} = $value->[1] ;
				$current_cell->{REF_FETCH_SUB} = $value->[2] ;
				$current_cell->{CACHE}         = 0 ;
				last
				} ;
				
			#----------------------
			# setting a value:
			#----------------------
			my $value_to_store = $value ; # do not modify $value as it is used again when storing ranges
			
			# check for range fillers
			if(/^Spreadsheet::Perl::RangeValues$/)
				{
				$value_to_store  = shift @$value  ;
				}
			else
				{
				if(/^Spreadsheet::Perl::RangeValuesSub$/)
					{
					$value_to_store = $value->[0]($self, $address, $current_address, @$value[1 .. (@$value - 1)]) ;
					}
				#else
					# store the value passed to STORE
				}
			
			if(exists $current_cell->{STORE_SUB})
				{
				if(exists $current_cell->{STORE_SUB_ARGS} && @{$current_cell->{STORE_SUB_ARGS}})
					{
					$current_cell->{STORE_SUB}->($self, $current_address, $value_to_store, @{$current_cell->{STORE_SUB_ARGS}}) ;
					}
				else
					{
					$current_cell->{STORE_SUB}->($self, $current_address, $value_to_store) ;
					}
				}
			else
				{
				# storing a simple value removes formulas

				delete $current_cell->{FORMULA} ;
				delete $current_cell->{PERL_FORMULA} ;
				delete $current_cell->{FETCH_SUB} ;
				delete $current_cell->{FETCH_SUB_ARGS} ;
				delete $current_cell->{GENERATED_FORMULA} ;
				delete $current_cell->{ANCHOR} ;

				$current_cell->{NEED_UPDATE} = 0 ;

				$current_cell->{VALUE} = $value_to_store ;
				}
				
			if(exists $current_cell->{REF_STORE_SUB})
				{
				$current_cell->{REF_STORE_SUB}->($self, $current_address, $value_to_store) ;
				}
			}
		
		if($self->{AUTOCALC} && exists $current_cell->{DEPENDENT} && $current_cell->{DEPENDENT})
			{
			$self->Recalculate() ;
			}

		if(exists $self->{DEBUG}{RECORD_STORE_ALL} || exists $self->{DEBUG}{RECORD_STORE}{$current_address})
			{	
			use Data::TreeDumper::Utils ;
			push @{$current_cell->{STORED_AT}}, Data::TreeDumper::Utils::get_caller_stack ;
			}
		}
	else
		{
		# not validated
		}
	}
}

*Set = \&STORE ;

#-------------------------------------------------------------------------------
1 ;

__END__
=head1 NAME

Spreadsheet::Perl::Store - implements Tie STORE

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

