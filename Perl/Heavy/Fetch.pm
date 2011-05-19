
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

sub FETCH 
{
my $self    = shift ;
my $address = shift;

my $attribute ;

if($address =~ /(.*)\.(.+)/)
	{
	$address = $1 ;
	$attribute = $2 ;
	}

#inter spreadsheet references
my $original_address = $address ;
my $ss_reference ;

my ($cell_or_range, $is_cell, $start_cell, $end_cell) = $self->CanonizeAddress($address) ;

($ss_reference, $address) = $self->GetSpreadsheetReference($cell_or_range) ;

if(defined $ss_reference)
	{
	if($ss_reference != $self)
		{
		if($self->{DEBUG}{FETCH_FROM_OTHER})
			{
			my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
			print $dh $self->GetName() . " Fetching from spreadsheet '$original_address'.\n" ;
			}
			
		#handle inter spreadsheet dependency tracking and formula references
		if(exists $self->{DEPENDENCY_STACK})
			{
			$ss_reference->{DEPENDENCY_STACK} = $self->{DEPENDENCY_STACK} ;
			$ss_reference->{DEPENDENCY_STACK_LEVEL} = $self->{DEPENDENCY_STACK_LEVEL} ;
			$ss_reference->{DEPENDENCY_STACK_NO_CACHE} = $self->{DEPENDENCY_STACK_NO_CACHE} ;
			}

		# all spreadsheets reference the same DEPENDENT_STACK
		$ss_reference->{DEPENDENT_STACK} = $self->{DEPENDENT_STACK} ;
		
		my $cell_value = $ss_reference->Get($address) ;
		
		# handle DEPENDENCY stack 
		
		# TODO: delete reference in ss_reference
		# be carefull in case C = B = A, and deleting C
		# means deleting A
		# have each get return a dependency stack instead

		#delete $ss_reference->{DEPENDENCY_STACK} ;
		#delete $ss_reference->{DEPENDENCY_STACK_NO_CACHE} ;

		$self->{DEPENDENCY_STACK_LEVEL} = $ss_reference->{DEPENDENCY_STACK_LEVEL} ;
		#delete $ss_reference->{DEPENDENCY_STACK_LEVEL} ;

		return($cell_value) ;
		}
	}
else
	{
	confess "Can't find Spreadsheet object for address '$address'.\n." ;
	}

if($self->{DEBUG}{FETCH})
	{
	my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
	my $name = $self->GetName() ;

	if($is_cell)
		{
		print $dh "Fetching cell '$name!$cell_or_range'\n" ;
		}
	else
		{
		print $dh "Fetching range '$name!$cell_or_range'\n" ;
		}
	}
	
if($is_cell)
	{
	my ($value, $evaluation_ok, $evaluation_type, $evaluation_data) ;
	
	# user defined trigger
	if(exists $self->{DEBUG}{FETCH_TRIGGER}{$start_cell})
		{
		if('CODE' eq ref $self->{DEBUG}{FETCH_TRIGGER}{$start_cell})
			{
			$self->{DEBUG}{FETCH_TRIGGER}{$start_cell}->($self, $start_cell, $attribute) ;
			}
		else
			{
			if(exists $self->{DEBUG}{FETCH_TRIGGER_HANDLER})
				{
				$self->{DEBUG}{FETCH_TRIGGER_HANDLER}->($self, $start_cell, $attribute) ;
				}
			else
				{
				my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
				print $dh "Fetching cell '$start_cell' (no fetch trigger handler set)\n" ;
				}
			}
		}
		
	if(exists $self->{CELLS}{$start_cell})
		{
		my $current_cell = $self->{CELLS}{$start_cell} ;
		
		if(defined $attribute)
			{
			$value = $current_cell->{$attribute} if(exists $current_cell->{$attribute}) ;
			}
		else
			{
			$current_cell->{FETCHED}++ if($self->{DEBUG}{FETCHED}) ;
				
			# circular dependency checking
			if(exists $current_cell->{CYCLIC_FLAG})
				{
				my $name = $self->GetName() ;

				if(exists $self->{DEPENDENCY_STACK})
					{
					my $level = '   ' x $self->{DEPENDENCY_STACK_LEVEL} ; 
					$self->{DEPENDENCY_STACK_LEVEL}++ ;

					push @{$self->{DEPENDENCY_STACK}}, "$level#CYCLIC $name!$start_cell" ;
					}

				my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
				
				push @{$self->{DEPENDENT_STACK}}, [$self, $start_cell] ;
				print $dh $self->DumpDependentStack("Cyclic dependency while fetching '$name!$start_cell'") ;
				
				my @dump ;

				for my $dependent (@{$self->{DEPENDENT_STACK}})
					{
					my ($spreadsheet, $address) = @$dependent ;
					push @dump, $spreadsheet->GetName() . '!' . $address ;
					}

				pop @{$self->{DEPENDENT_STACK}} ;

				# Todo: is there some cleanup to do before calling die?
				die bless {cycle => \@dump}, 'Cyclic dependency' ;
				}
			else
				{
				$current_cell->{CYCLIC_FLAG}++ ;
				}
		
			if(exists $self->{DEPENDENCY_STACK})
				{
				my $level = '   ' x $self->{DEPENDENCY_STACK_LEVEL} ; 
				$self->{DEPENDENCY_STACK_LEVEL}++ ;

				my $name = $self->GetName() ;

				my $formula = exists $current_cell->{PERL_FORMULA}
									? ': ' . $current_cell->{GENERATED_FORMULA}
									: exists $current_cell->{FORMULA}
										? ': ' . $current_cell->{GENERATED_FORMULA}
										: '' ;

				push @{$self->{DEPENDENCY_STACK}}, "$level$name!$start_cell$formula" ;
				}

			$self->FindDependent($current_cell, $start_cell) ;
			push @{$self->{DEPENDENT_STACK}}, [$self, $start_cell] ;
			
			if($self->{DEBUG}{DEPENDENT_STACK_ALL} || $self->{DEBUG}{DEPENDENT_STACK}{$start_cell})
				{
				my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
				print $dh $self->DumpDependentStack("Fetching '" . $self->GetName() . "!$start_cell'") ;
				}
				
			# formula directly set into cells must get "compiled"
			# IE, when data is directly loaded from external file
			if(exists $current_cell->{PERL_FORMULA} && ! exists $current_cell->{FETCH_SUB})
				{
				die "case to be handled!\n" ;
				}
			else
				{
				if(exists $current_cell->{FORMULA} && ! exists $current_cell->{FETCH_SUB})
					{
					die "case to be handled!\n" ;
					}
				}
				
			if(exists $current_cell->{FETCH_SUB}) # formula or fetch callback
				{
				$self->initial_value_from_perl_scalar($start_cell, $current_cell)  if(exists $current_cell->{REF_FETCH_SUB}) ;
				
				if
					(
					$current_cell->{NEED_UPDATE} 
					|| ! exists $current_cell->{NEED_UPDATE} 
					|| ! exists $current_cell->{VALUE}
					|| exists $self->{DEPENDENCY_STACK_NO_CACHE} 
					)
					{
					if($self->{DEBUG}{FETCH_SUB})
						{
						my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
						my $ss_name = $self->GetName() ;
						
						print $dh "Running Sub @ '$ss_name!$start_cell'" ;
						
						if(exists $current_cell->{FORMULA})
							{
							print $dh " formula: $current_cell->{FORMULA}[1]" ;
							}
							
						if(exists $current_cell->{PERL_FORMULA})
							{
							print $dh " formula: $current_cell->{PERL_FORMULA}[1]" ;
							}
							
						print $dh " defined at '@{$current_cell->{DEFINED_AT}}'" if(exists $current_cell->{DEFINED_AT}) ;
						print $dh "\n" ;
						}
#TODO!!!!  next section should be in eval and formula should generate an exception. this is needed so a failed formula doesn't update NEED_UPDATE state nor saves the erroneous value trought STORE_ON_FETCH or a perl scalar. Although I am not sure. Maybe it is better to store the erroneous values if they are descriptive enought so it is clear that the values are not in synch.	
					if(exists $current_cell->{FETCH_SUB_ARGS} && @{$current_cell->{FETCH_SUB_ARGS}})
						{
						($value, $evaluation_ok, $evaluation_type, $evaluation_data)
							= ($current_cell->{FETCH_SUB})->($self, $start_cell, @{$current_cell->{FETCH_SUB_ARGS}}) ;
						}
						
					else
						{
						($value, $evaluation_ok, $evaluation_type, $evaluation_data)
							= ($current_cell->{FETCH_SUB})->($self, $start_cell) ;
						}
						
					if(exists $current_cell->{REF_STORE_SUB} && exists $current_cell->{STORE_ON_FETCH})
						{
						$current_cell->{REF_STORE_SUB}->($self, $start_cell, $value) ;
						}
						
					if(exists $current_cell->{STORE_SUB} && exists $current_cell->{STORE_ON_FETCH})
						{
						if(exists $current_cell->{STORE_SUB_ARGS} && @{$current_cell->{STORE_SUB_ARGS}})
							{
							$current_cell->{STORE_SUB}->($self, $start_cell, $value, @{$current_cell->{STORE_SUB_ARGS}}) ;
							}
						else
							{
							$current_cell->{STORE_SUB}->($self, $start_cell, $value) ;
							}
						}
						
					if($self->{DEBUG}{PRINT_FORMULA_ERROR} && $evaluation_ok == 0)
						{
						$value .= " ($evaluation_type)" ;
						}

					$current_cell->{EVAL_TYPE} = $evaluation_type ;
					$current_cell->{EVAL_OK} = $evaluation_ok ;
					$current_cell->{EVAL_DATA} = $evaluation_data ;
					
					# handle caching
					if((! $self->{CACHE}) || (exists $current_cell->{CACHE} && (! $current_cell->{CACHE})))
						{
						delete $current_cell->{VALUE} ;
						$current_cell->{NEED_UPDATE} = 1 ;
						}
					else
						{
						$current_cell->{VALUE} = $value ;
						$current_cell->{NEED_UPDATE} = 0 ;
						}

					if(@{$self->{DEPENDENT_STACK}} != 1)
						{
						# catch exception at cell that started computation
						unless ($evaluation_ok)
							{
							#Todo: cleanup automatically
							$self->{DEPENDENCY_STACK_LEVEL}-- if exists $self->{DEPENDENCY_STACK_LEVEL} ;
					
							pop @{$self->{DEPENDENT_STACK}} ;
							delete $current_cell->{CYCLIC_FLAG} ;

							die bless {spreadsheet => $self, cell => $start_cell}, 'Invalid dependency cell' ;
							}
						}
					}
				else
					{
					$value = $current_cell->{VALUE} ;
					}
				}
			else
				{
				if(exists $current_cell->{REF_FETCH_SUB})
					{
					#fetch value from reference
					if(exists $self->{DEBUG}{FETCH_TRIGGER}{$start_cell})
						{
						my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
						print $dh "  => Fetching cell '$start_cell' value from scalar reference.\n" ;
						}
					
					$current_cell->{VALUE} = $current_cell->{REF_FETCH_SUB}->($self, $start_cell) ;
					}
					
				# Todo: shouldn't we handle cache here too?!
				if(exists $current_cell->{VALUE})
					{
					$value = $current_cell->{VALUE} ;
					}
				else
					{
					$value = undef ;
					}
				}
			$self->{DEPENDENCY_STACK_LEVEL}-- if exists $self->{DEPENDENCY_STACK_LEVEL} ;
				
			pop @{$self->{DEPENDENT_STACK}} ;
			delete $current_cell->{CYCLIC_FLAG} ;
			}
		}
	else
		{
		# cell has never been accessed before
		# not even to set a formula or a dependency list in it
		if(exists $self->{DEPENDENCY_STACK})
			{
			my $level = '   ' x $self->{DEPENDENCY_STACK_LEVEL} ; 
			my $name = $self->GetName() . '!' ;

			push @{$self->{DEPENDENCY_STACK}}, "$level$name$start_cell" ;
			}

		if(@{$self->{DEPENDENT_STACK}})
			{
			$self->{CELLS}{$start_cell} = {} ; # create the cell to hold the dependent
			$self->FindDependent($self->{CELLS}{$start_cell}, $start_cell) ;
		
			if($self->{DEBUG}{DEPENDENT_STACK_ALL} || $self->{DEBUG}{DEPENDENT_STACK}{$start_cell})
				{
				push @{$self->{DEPENDENT_STACK}}, [$self, $start_cell] ;
				
				my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
				print $dh $self->DumpDependentStack("Fetching '" . $self->GetName() . "!$start_cell' (virtual cell)") ;
				
				pop @{$self->{DEPENDENT_STACK}} ;
				}
			}
			
		# handle headers and default values
		my ($column, $row) = ConvertAdressToNumeric($start_cell) ;
		
		if($row == 0)
			{
			$value = ToAA($column) ;
			}
		else
			{
			if($column == 0)
				{
				$value = $self->{MESSAGE}{ROW_PREFIX} . $row ;
				}
			else
				{
				$value = undef ;
				}
			}
		}

	if($self->{DEBUG}{FETCH_VALUE})
		{
		my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
		print $dh "\t value: $value\n" ;
		}
		
	return($value) ;
	}
else
	{
	# range requested
	
	my @values ;

	for my $current_address ($self->GetAddressList($address))
		{
		push @values, $self->Get($current_address) ;
		}

	if($self->{DEBUG}{FETCH})
		{
		my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
		print $dh "END: Fetching range '$cell_or_range'\n" ;
		}

		
	return \@values ;
	}
}

*Get = \&FETCH ;

#-------------------------------------------------------------------------------

sub initial_value_from_perl_scalar
{
# note that the scalar fetch mechanism is removed after the call to this sub

my ($self, $cell_address, $current_cell) = @_ ;

if(exists $current_cell->{REF_FETCH_SUB})
	{
	# value from reference will be shdowed by formula
	if(exists $self->{DEBUG}{FETCH_TRIGGER}{$cell_address})
		{
		my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
		print $dh "  => Cell '$cell_address' value from scalar reference is shadowed by formula.\n" ;
		}
	
	$current_cell->{VALUE} = '# shadowed by formula' ;
	
	# the value from the scalar can be fetched with
	# $current_cell->{REF_FETCH_SUB}->($self, $cell_address) ;
	
	delete $current_cell->{REF_FETCH_SUB} ;
	delete $current_cell->{CACHE} ;
	$current_cell->{STORE_ON_FETCH}++ ;
	}
}
			

#-------------------------------------------------------------------------------
1 ;

__END__
=head1 NAME

Spreadsheet::Perl::Fetch - Implements Tie FETCH

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

