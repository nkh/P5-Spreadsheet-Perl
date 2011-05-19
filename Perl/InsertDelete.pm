
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

our $VERSION = '0.03' ;

#-------------------------------------------------------------------------------

# NOTE: When updating cell dependencies and formulas, the dependencies must be updated first
# as the update of the formulas triggers the marking of the cell dependencies for update
# which cell to update depends on the dependent list, which must be updated first. if the triggering
# mechanism finds a dependent that does not exist, and exception is thrown. Be careful to update
# ALL the dependents before updating any formula
# triggering the cells should be done only when a cell is deleted, insertion has no effect on single cell
# dependency but may have on range dependency

sub InsertRows
{
my ($self, $start_row, $number_of_rows_to_insert) = @_ ;

confess "Invalid row '$start_row'\n" unless $start_row =~ /^\s*\d+\s*$/ ;

if($self->{DEBUG}{OFFSET_ADDRESS})
	{
	print "-------------------------------------------------------------------\n" ;
	print "Inserting rows in " . $self->GetName() . "\n" ;
	print "-------------------------------------------------------------------\n" ;
	}

my (%moved_cell_list, %not_moved_cell_list) ;

for my $cell_address ($self->GetCellList())
	{
	# get all the cells for the rows under the $start_row
	my ($column, $row) = $cell_address =~ /([A-Z]+)(\d+)/ ;

	if( $row >= $start_row)
		{
		push @{$moved_cell_list{$row}}, $cell_address ;
		}
	else
		{
		push @{$not_moved_cell_list{$row}}, $cell_address ;
		}
	}

for my $row (reverse sort keys %moved_cell_list)
	{
	for my $cell_address (@{$moved_cell_list{$row}})
		{
		my $new_address = $self->OffsetAddress($cell_address, 0, $number_of_rows_to_insert) ; 
		
		if(exists $self->{CELLS}{$cell_address}{DEPENDENT})
			{
			$self->OffsetDependents($self, $cell_address, 0, 0, $start_row, $number_of_rows_to_insert, "A$start_row:AAAA9999") ;
			}

		if(exists $self->{CELLS}{$cell_address}{GENERATED_FORMULA})
			{
			$self->OffsetFormula($cell_address, 0, 0, $start_row, $number_of_rows_to_insert, "A$start_row:AAAA9999", $self) ;
			}

		$self->{CELLS}{$new_address} = $self->{CELLS}{$cell_address} ;
		delete $self->{CELLS}{$cell_address} ;
		}
	}

# note, the cells don't have to be update in a specific order
# we keep the same order as moved cells to create the illusion
# of order
for my $row (reverse sort keys %not_moved_cell_list)
	{
	for my $cell_address (@{$not_moved_cell_list{$row}})
		{
		if(exists $self->{CELLS}{$cell_address}{DEPENDENT})
			{
			$self->OffsetDependents($self, $cell_address, 0, 0, $start_row, $number_of_rows_to_insert, "A$start_row:AAAA9999") ;
			}
			
		if(exists $self->{CELLS}{$cell_address}{GENERATED_FORMULA})
			{
			$self->OffsetFormula($cell_address, 0, 0, $start_row, $number_of_rows_to_insert, "A$start_row:AAAA9999", $self) ;
			}

		}
	}

$self->UpdateOtherSpreadsheet(0, 0, $start_row, $number_of_rows_to_insert, "A$start_row:AAAA9999") ;

for my $row_header (reverse SortCells grep {/^@/} $self->GetCellHeaderList())
	{
	my ($row_index) = $row_header =~ /^@(.+)/ ;
	if($row_index >= $start_row)
		{
		my $new_row = $row_index + $number_of_rows_to_insert ;
		$self->{CELLS}{"\@$new_row"} = $self->{CELLS}{$row_header} ;
		delete $self->{CELLS}{$row_header} ;	
		}
	}
}


sub UpdateOtherSpreadsheet
{
my ($self, $column_offset, $number_of_columns_to_insert, $row_offset, $number_of_rows_to_insert, $range) = @_ ;

if($self->{DEBUG}{OFFSET_ADDRESS})
	{
	print "-------------------------------------------------------------------\n" ;
	print "Updating other spreadsheets after modification to " . $self->GetName() . "\n" ;
	print "-------------------------------------------------------------------\n" ;
	}

for my $other_spreadsheet (values  %{$self->{OTHER_SPREADSHEETS}})
	{
	for my $other_cell ($other_spreadsheet->GetCellList())
		{
		if(exists $other_spreadsheet->{CELLS}{$other_cell}{DEPENDENT})
			{
			$other_spreadsheet->OffsetDependents($self, $other_cell, $column_offset, $number_of_columns_to_insert, $row_offset, $number_of_rows_to_insert, $range) ;
			}
		}
			
	for my $other_cell ($other_spreadsheet->GetCellList())
		{
		if(exists $other_spreadsheet->{CELLS}{$other_cell}{GENERATED_FORMULA})
			{
			$other_spreadsheet->OffsetFormula($other_cell, $column_offset, $number_of_columns_to_insert, $row_offset, $number_of_rows_to_insert, $range, $self) ;
			}
		}
	}
	
print "-------------------------------------------------------------------\n\n" if($self->{DEBUG}{OFFSET_ADDRESS}) ;
}

#-------------------------------------------------------------------------------------------------------

sub InsertColumns
{
my ($self, $start_column, $number_of_columns_to_insert) = @_ ;

confess "Invalid '$start_column'\n" unless $start_column =~ /^\s*[A-Z]{1,4}\s*$/ ;

if($self->{DEBUG}{OFFSET_ADDRESS})
	{
	print "-------------------------------------------------------------------\n" ;
	print "Inserting columns in " . $self->GetName() . "\n" ;
	print "-------------------------------------------------------------------\n" ;
	}

my (%moved_cell_list, %not_moved_cell_list) ;

for my $cell_address ($self->GetCellList())
	{
	# get all the cells for the rows under the $start_row
	my ($column, $row) = $cell_address =~ /([A-Z]+)(\d+)/ ;

	my $column_index = FromAA($column) ;
	my $start_column_index = FromAA($start_column) ;

	if( $column_index >= $start_column_index)
		{
		push @{$moved_cell_list{$column_index}}, $cell_address ;
		}
	else
		{
		push @{$not_moved_cell_list{$column_index}}, $cell_address ;
		}
	}

for my $column_index (reverse sort keys %moved_cell_list)
	{
	for my $cell_address (@{$moved_cell_list{$column_index}})
		{
		my $new_address = $self->OffsetAddress($cell_address, $number_of_columns_to_insert, 0) ; 
		
		if(exists $self->{CELLS}{$cell_address}{DEPENDENT})
			{
			$self->OffsetDependents($self, $cell_address, $start_column, $number_of_columns_to_insert, 0, 0, "${start_column}1:AAAA9999") ;
			}

		if(exists $self->{CELLS}{$cell_address}{GENERATED_FORMULA})
			{
			$self->OffsetFormula($cell_address, $start_column, $number_of_columns_to_insert, 0, 0, "${start_column}1:AAAA9999", $self) ;
			}

		$self->{CELLS}{$new_address} = $self->{CELLS}{$cell_address} ;
		delete $self->{CELLS}{$cell_address} ;
		}
	}

# note, the cells don't have to be update in a specific order
# we keep the same order as moved cells to create the illusion
# of order
for my $column_index (reverse sort keys %not_moved_cell_list)
	{
	for my $cell_address (@{$not_moved_cell_list{$column_index}})
		{
		if(exists $self->{CELLS}{$cell_address}{DEPENDENT})
			{
			$self->OffsetDependents($self, $cell_address, $start_column, $number_of_columns_to_insert, 0, 0, "${start_column}1:AAAA9999") ;
			}
			
		if(exists $self->{CELLS}{$cell_address}{GENERATED_FORMULA})
			{
			$self->OffsetFormula($cell_address, $start_column, $number_of_columns_to_insert, 0, 0, "${start_column}1:AAAA9999", $self) ;
			}
		}
	}

$self->UpdateOtherSpreadsheet($start_column, $number_of_columns_to_insert, 0, 0, "${start_column}1:AAAA9999") ;


my $start_column_index = FromAA($start_column) ;

for my $column_header (reverse SortCells grep {/^[A-Z]+0$/} $self->GetCellHeaderList())
	{
	my ($column_index) = $column_header =~ /^([A-Z]+)0$/ ;
	$column_index = FromAA($column_index) ;

	if($column_index >= $start_column_index)
		{
		my $new_column = $column_index + $number_of_columns_to_insert ;
		$new_column = ToAA($new_column) ;

		$self->{CELLS}{"${new_column}0"} = $self->{CELLS}{$column_header} ;
		delete $self->{CELLS}{$column_header} ;	
		}
	}
}

#-------------------------------------------------------------------------------

sub OffsetFormula
{
my 
	(
	$self, $cell_address,
	$start_column, $columns_to_insert,
	$start_row, $rows_to_insert,
	$range, $dependency_spreadsheet
	)  = @_ ;

print 'Updating ' . $self->GetName() . "!$cell_address formula\n" if($self->{DEBUG}{OFFSET_ADDRESS}) ;

return unless exists $self->{CELLS}{$cell_address}{GENERATED_FORMULA} ;

my $formula = $self->{CELLS}{$cell_address}{GENERATED_FORMULA} ; 
$formula =~ s/(([A-Za-z_0-9]+!)?\[?[A-Z]+\]?\[?[0-9]+\]?(:\[?[A-Z]+\]?\[?[0-9]+\]?)?)/$self->OffsetAddress($1, $columns_to_insert, $rows_to_insert, $range, $dependency_spreadsheet)/eg ;

$self->Set($cell_address, PF($formula)) ;
}


#-------------------------------------------------------------------------------

sub OffsetDependents
{
my 
	(
	$self, $dependent_spreadsheet,
	$cell_address,
	$start_column, $columns_to_insert,
	$start_row, $rows_to_insert,
	$range
	)  = @_ ;

print 'Updating ' . $self->GetName() . "!$cell_address dependents\n" if($self->{DEBUG}{OFFSET_ADDRESS}) ;

return unless exists $self->{CELLS}{$cell_address}{DEPENDENT} ;

my $dependents = $self->{CELLS}{$cell_address}{DEPENDENT} ; 

my @new_dependents ;

for my $dependent_name (keys %{$dependents})
	{
	my $dependent = $dependents->{$dependent_name} ;

	my ($spreadsheet, $cell_name) = @{$dependent->{DEPENDENT_DATA}} ;
	
	next if $spreadsheet != $dependent_spreadsheet ; 
	
	my $spreadsheet_name = $spreadsheet->GetName() ;

	my $new_cell_name = $self->OffsetAddress("$spreadsheet_name!$cell_name", $columns_to_insert, $rows_to_insert, $range, $dependent_spreadsheet) ;

	$new_cell_name =~ s/[A-Za-z_0-9]+!// ;
	
	$dependent->{DEPENDENT_DATA}[1] = $new_cell_name ;
	push @new_dependents, $dependent ;

	delete $dependents->{$dependent_name} ;
	}

for my $dependent (@new_dependents)
	{
	my $dependent_name = $dependent->{DEPENDENT_DATA}[0]->GetName() . '!' . $dependent->{DEPENDENT_DATA}[1] ;
	$dependents->{$dependent_name} = $dependent ;
	}
}

#-------------------------------------------------------------------------------

sub DeleteDependents
{
my ($self, $cell_address, $range, $dependent_spreadsheet) = @_ ;

# delete any dependent that is within the range

return unless exists $self->{CELLS}{$cell_address}{DEPENDENT} ;

my $dependents = $self->{CELLS}{$cell_address}{DEPENDENT} ; 

for my $dependent_name (keys %{$dependents})
	{
	my $dependent = $dependents->{$dependent_name} ;
	my ($spreadsheet, $cell_name) = @{$dependent->{DEPENDENT_DATA}} ;
	next unless $spreadsheet == $dependent_spreadsheet ; 

	if($self->is_within_range($cell_name, $range))
		{
		print "Deleting dependents $dependent_name at " . $self->GetName() . "!$cell_address\n" if($self->{DEBUG}{OFFSET_ADDRESS}) ;
		
		delete $dependents->{$dependent_name} ;
		}
	}
}

sub DeleteDependentsInOtherSpreadsheets
{
my ($self, $range) = @_ ;
	
for my $other_spreadsheet (values  %{$self->{OTHER_SPREADSHEETS}})
	{
	for my $other_cell ($other_spreadsheet->GetCellList())
		{
		if(exists $other_spreadsheet->{CELLS}{$other_cell}{DEPENDENT})
			{
			$other_spreadsheet->DeleteDependents($other_cell, $range, $self) ;
			}
		}
	}
}

#-------------------------------------------------------------------------------

sub UpdateOtherSpreadsheetAfterDelete
{
my ($self, $column_offset, $number_of_columns_to_insert, $row_offset, $number_of_rows_to_insert, $deleted_range, $offset_range) = @_ ;

if($self->{DEBUG}{OFFSET_ADDRESS})
	{
	print "-------------------------------------------------------------------\n" ;
	print "Updating other spreadsheets after modification to " . $self->GetName() . "\n" ;
	print "-------------------------------------------------------------------\n" ;
	}

$self->DeleteDependentsInOtherSpreadsheets($deleted_range) ;

for my $other_spreadsheet (values  %{$self->{OTHER_SPREADSHEETS}})
	{
	for my $other_cell ($other_spreadsheet->GetCellList())
		{
		if(exists $other_spreadsheet->{CELLS}{$other_cell}{DEPENDENT})
			{
			$other_spreadsheet->OffsetDependents($self, $other_cell, $column_offset, $number_of_columns_to_insert, $row_offset, $number_of_rows_to_insert, $offset_range) ;
			}
		}
			
	for my $other_cell ($other_spreadsheet->GetCellList())
		{
		if(exists $other_spreadsheet->{CELLS}{$other_cell}{GENERATED_FORMULA})
			{
			if($other_spreadsheet->FormulaReferenceRange($self, $other_cell, $deleted_range)) 
				{
				$other_spreadsheet->Set($other_cell, PF("'#REF [d]'")) ;
				}	
			else
				{
				$other_spreadsheet->OffsetFormula($other_cell, $column_offset, $number_of_columns_to_insert, $row_offset, $number_of_rows_to_insert, $offset_range, $self) ;
				}
			}
		}
	}
	
print "-------------------------------------------------------------------\n\n" if($self->{DEBUG}{OFFSET_ADDRESS}) ;
}

#-------------------------------------------------------------------------------

sub DeleteColumns
{
my ($self, $start_column, $number_of_columns_to_delete) = @_ ;

confess "Invalid '$start_column'\n" unless $start_column =~ /^\s*[A-Z]{1,4}\s*$/ ;

if($self->{DEBUG}{OFFSET_ADDRESS})
	{
	print "-------------------------------------------------------------------\n" ;
	print "Deleting columns in " . $self->GetName() . "\n" ;
	print "-------------------------------------------------------------------\n" ;
	}

my $start_column_index = FromAA($start_column) ;
my $end_column = ToAA($start_column_index + $number_of_columns_to_delete - 1) ;

my (%removed_cell_list, %moved_cell_list, %not_moved_cell_list) ;

for my $cell_address ($self->GetCellList())
	{
	# get all the cells for the rows under the $start_row
	my ($column, $row) = $cell_address =~ /([A-Z]+)(\d+)/ ;

	my $column_index = FromAA($column) ;
	my $start_column_index = FromAA($start_column) ;
	
	if( $column_index >= $start_column_index)
		{
		if ($column_index < $start_column_index + $number_of_columns_to_delete)
			{
			push @{$removed_cell_list{$column_index}}, $cell_address ;
			}
		else
			{
			push @{$moved_cell_list{$column_index}}, $cell_address ;
			}
		}
	else
		{
		push @{$not_moved_cell_list{$column_index}}, $cell_address ;
		}
	}

for my $column_index (keys %removed_cell_list)
	{
	for my $cell_address (@{$removed_cell_list{$column_index}})
		{
		$self->DELETE($cell_address) ; # DELETE calls user defined callbacks
		}
	}

for my $column_index (sort keys %moved_cell_list)
	{
	for my $cell_address (@{$moved_cell_list{$column_index}})
		{
		if(exists $self->{CELLS}{$cell_address}{DEPENDENT})
			{
			$self->DeleteDependents($cell_address, "${start_column}1:${end_column}9999", $self) ;
			$self->OffsetDependents($self, $cell_address, $start_column, - $number_of_columns_to_delete, 0, 0, "${start_column}1:AAAA9999") ;
			}

		if(exists $self->{CELLS}{$cell_address}{GENERATED_FORMULA})
			{
			if($self->FormulaReferenceRange($self, $cell_address, "${start_column}1:${end_column}9999")) 
				{
				$self->Set($cell_address, PF("'#REF [dc]'")) ;
				}	
			else
				{
				$self->OffsetFormula($cell_address, $start_column, - $number_of_columns_to_delete, 0, 0, "${start_column}1:AAAA9999", $self) ;
				}
			}

		my $new_address = $self->OffsetAddress($cell_address, - $number_of_columns_to_delete, 0) ;

		$self->{CELLS}{$new_address} = $self->{CELLS}{$cell_address} ;
		delete $self->{CELLS}{$cell_address} ;
		}
	}

# note, the cells don't have to be update in a specific order
# we keep the same order as moved cells to create the illusion
# of order
for my $column_index (reverse sort keys %not_moved_cell_list)
	{
	for my $cell_address (@{$not_moved_cell_list{$column_index}})
		{
		# TODO GENERATED_FORMULA exists only after the cell has been 
		# compiled. Is there a case where DeleteColumns could be called 
		# before the formula generation?
		
		if(exists $self->{CELLS}{$cell_address}{DEPENDENT})
			{
			$self->DeleteDependents($cell_address, "${start_column}1:${end_column}9999", $self) ;
			$self->OffsetDependents($self, $cell_address, $start_column, - $number_of_columns_to_delete, 0, 0, "${start_column}1:AAAA9999") ;
			}
			
		if(exists $self->{CELLS}{$cell_address}{GENERATED_FORMULA})
			{
			if($self->FormulaReferenceRange($self, $cell_address, "${start_column}1:${end_column}9999")) 
				{
				$self->Set($cell_address, PF("'#REF [dc]'")) ;
				}	
			else
				{
				$self->OffsetFormula($cell_address, $start_column, - $number_of_columns_to_delete, 0, 0, "${start_column}1:AAAA9999", $self) ;
				}
			}
		}
	}

$self->UpdateOtherSpreadsheetAfterDelete($start_column, - $number_of_columns_to_delete, 0, 0, "${start_column}1:${end_column}9999", "${start_column}1:AAAA9999") ;

for my $column_header (SortCells grep {!/^@/} $self->GetCellHeaderList())
	{
	my ($column_index) = $column_header =~ /^([A-Z@]+)0$/ ;
	$column_index = FromAA($column_index) ;

	if($column_index >= $start_column_index)
		{
		if ($column_index < $start_column_index + $number_of_columns_to_delete)
			{
			$self->DELETE($column_header) ;	
			}
		else
			{
			my $new_column = $column_index - $number_of_columns_to_delete ;
			$new_column = ToAA($new_column) ;

			$self->{CELLS}{"${new_column}0"} = $self->{CELLS}{$column_header} ;
			delete $self->{CELLS}{$column_header} ;	
			}
		}
	}
}

#-------------------------------------------------------------------------------

sub DeleteRows
{
my ($self, $start_row, $number_of_rows_to_delete) = @_ ;

confess "Invalid '$start_row'\n" unless $start_row =~ /^\s*\d+\s*$/ ;

if($self->{DEBUG}{OFFSET_ADDRESS})
	{
	print "-------------------------------------------------------------------\n" ;
	print "Deleting rows in " . $self->GetName() . "\n" ;
	print "-------------------------------------------------------------------\n" ;
	}

my $end_row = $start_row + $number_of_rows_to_delete - 1 ;

my (%removed_cell_list, %moved_cell_list, %not_moved_cell_list) ;

for my $cell_address ($self->GetCellList())
	{
	# get all the cells for the rows under the $start_row
	my ($column, $row) = $cell_address =~ /([A-Z]+)(\d+)/ ;

	if( $row >= $start_row)
		{
		if ($row < $start_row + $number_of_rows_to_delete)
			{
			push @{$removed_cell_list{$row}}, $cell_address ;
			}
		else
			{
			push @{$moved_cell_list{$row}}, $cell_address ;
			}
		}
	else
		{
		push @{$not_moved_cell_list{$row}}, $cell_address ;
		}
	}

for my $row (keys %removed_cell_list)
	{
	for my $cell_address (@{$removed_cell_list{$row}})
		{
		$self->DELETE($cell_address) ; # DELETE would call the appropriate callback
		}
	}

for my $row (sort keys %moved_cell_list)
	{
	for my $cell_address (@{$moved_cell_list{$row}})
		{
		if(exists $self->{CELLS}{$cell_address}{DEPENDENT})
			{
			$self->DeleteDependents($cell_address,  "A${start_row}:AAAA${end_row}", $self) ;
			$self->OffsetDependents($self, $cell_address, 0, 0, $start_row, - $number_of_rows_to_delete, "A${start_row}:AAAA9999") ;
			}

		if(exists $self->{CELLS}{$cell_address}{GENERATED_FORMULA})
			{
			if($self->FormulaReferenceRange($self, $cell_address, "A${start_row}:AAAA${end_row}")) 
				{
				$self->Set($cell_address, PF("'#REF [dr]'")) ;
				}	
			else
				{
				$self->OffsetFormula($cell_address, 0, 0, $start_row, - $number_of_rows_to_delete, "A${start_row}:AAAA9999", $self) ;
				}
			}

		my $new_address = $self->OffsetAddress($cell_address, 0, - $number_of_rows_to_delete) ;

		$self->{CELLS}{$new_address} = $self->{CELLS}{$cell_address} ;
		delete $self->{CELLS}{$cell_address} ;
		}
	}

# note, the cells don't have to be update in a specific order
# we keep the same order as moved cells to create the illusion
# of order
for my $row (reverse sort keys %not_moved_cell_list)
	{
	for my $cell_address (@{$not_moved_cell_list{$row}})
		{
		# TODO GENERATED_FORMULA exists only after the cell has been 
		# compiled. Is there a case where DeleteColumns could be called 
		# before the formula generation?
		
		if(exists $self->{CELLS}{$cell_address}{DEPENDENT})
			{
			$self->DeleteDependents($cell_address,  "A${start_row}:AAAA${end_row}", $self) ;
			$self->OffsetDependents($self, $cell_address, 0, 0, $start_row, - $number_of_rows_to_delete, "A${start_row}:AAAA9999") ;
			}
			
		if(exists $self->{CELLS}{$cell_address}{GENERATED_FORMULA})
			{
			if($self->FormulaReferenceRange($self, $cell_address, "A${start_row}:AAAA${end_row}")) 
				{
				$self->Set($cell_address, PF("'#REF [dr]'")) ;
				}	
			else
				{
				$self->OffsetFormula($cell_address, 0, 0, $start_row, - $number_of_rows_to_delete, "A${start_row}:AAAA9999", $self) ;
				}
			}
		}
	}

$self->UpdateOtherSpreadsheetAfterDelete(0, 0, $start_row, - $number_of_rows_to_delete, "A${start_row}1:AAAA${end_row}", "A${start_row}1:AAAA$9999") ;

for my $row_header (sort grep {/^@/} $self->GetCellHeaderList())
	{
	my ($row) = $row_header =~ /(\d+)$/ ;

	if($row>= $start_row)
		{
		if ($row < $start_row + $number_of_rows_to_delete)
			{
			$self->DELETE($row_header) ;	
			}
		else
			{
			my $new_row = $row - $number_of_rows_to_delete ;

			$self->{CELLS}{"\@${new_row}"} = $self->{CELLS}{$row_header} ;
			delete $self->{CELLS}{$row_header} ;	
			}
		}
	}
}

#-------------------------------------------------------------------------------

sub FormulaReferenceRange
{
my ($self, $spreadsheet_to_check, $cell_address, $range) = @_ ;

my $reference_range = 0 ;
my $formula = '' ;

if(exists $self->{CELLS}{$cell_address}{GENERATED_FORMULA})
	{
	my $spreadsheet_to_check_name = $spreadsheet_to_check->GetName() ;
	
	my ($rcs, $rrs, $rce, $rre) = $range =~/([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)/ ;

	unless(defined $rcs && defined $rrs && defined $rce && defined $rre)
		{
		confess "Invalid range '$range'\n" ;
		}
	
	($rcs, $rce) = (FromAA($rcs), FromAA($rce)) ;
	if($rcs > $rce)
		{
		if($rrs > $rre)
			{
			($rcs, $rrs, $rce, $rre) = ($rce, $rre, $rcs, $rrs) ;
			}
		else
			{
			($rcs, $rce) = ($rce, $rcs) ;
			}
		}
	else
		{
		if($rrs > $rre)
			{
			($rrs, $rre) = ($rre, $rrs) ;
			}
		#else 
		#	range in right order
		}


	$formula = $self->{CELLS}{$cell_address}{GENERATED_FORMULA} ;
	my ($fcs, $frs, $fce, $fre) ; 

	while($formula =~ /(([A-Za-z_0-9]+!)?(([A-Z]+)([0-9]+))(:([A-Z]+)([0-9]+))?)/g)
		{
		if(defined $2) # spreadheet name in formula
			{
			next unless ($spreadsheet_to_check_name . '!') eq $2 ;
			}
		else
			{
			next unless $spreadsheet_to_check == $self ;
			}

		if(defined $6)
			{
			# range
			($fcs, $frs, $fce, $fre) = ($4, $5, $7, $8) ;

			($fcs, $fce) = (FromAA($fcs), FromAA($fce)) ;

			if($fcs > $fce)
				{
				if($frs > $fre)
					{
					($fcs, $frs, $fce, $fre) = ($fce, $fre, $fcs, $frs) ;
					}
				else
					{
					($fcs, $fce) = ($fce, $fcs) ;
					}
				}
			else
				{
				if($frs > $fre)
					{
					($frs, $fre) = ($fre, $frs) ;
					}
				#else 
				#	range in right order
				}
			}
		else
			{
			($fcs, $frs) = ($4, $5) ;
			($fcs) = FromAA($fcs) ;

			($fce, $fre) = ($fcs, $frs) ; 
			}
			
		if
			(
			$fcs > $rce
			|| $fce < $rcs
			|| $frs > $rre
			|| $fre < $rrs
			)
			{
			# does not reference range
			}
		else
			{
			$reference_range++ ; # references range
			last ;
			}
		}
	}
#else
	# no formula
	
#~ print "FormulaReferenceRange: $cell_address, $formula, $range, $reference_range\n" ;

return $reference_range ,
}

#-------------------------------------------------------------------------------

1 ;

__END__

=head1 NAME

Spreadsheet::Perl::InsertDelete - Columns and rows insertion and deletion

=head1 SYNOPSIS

Part of Spreadsheet::Perl.

=head1 AUTHOR

Khemir Nadim ibn Hamouda. <nadim@khemir.net>

  Copyright (c) 2011 Nadim Ibn Hamouda el Khemir. All rights
  reserved.  This program is free software; you can redis-
  tribute it and/or modify it under the same terms as Perl
  itself.
  
If you find any value in this module, mail me!  All hints, tips, flames and wishes
are welcome at <nadim@khemir.net>.

=cut
