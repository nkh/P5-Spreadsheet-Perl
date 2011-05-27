
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

our $VERSION = '0.01' ;


=head1 NAME

 Spreadsheet:Perl::Heavy::Tie - Implementation of the tie mechanisms

=head1 SUBROUTINES/METHODS

Subroutines that are not part of the public interface are marked with [p].

=cut


#-------------------------------------------------------------------------------

sub GetSpreadsheetDefaultData
{

=head2 [p] GetSpreadsheetDefaultData()

I<Arguments> - None

I<Returns> 

=over 2

=item * list of the default data

=back

=cut

return 
	(
	  NAME                => undef
	, CACHE               => 1
	, AUTOCALC            => 0
	, OTHER_SPREADSHEETS  => {}
	, DEBUG               => {
								ERROR_HANDLE => \*STDERR,
								PRINT_FORMULA_ERROR => 1,
							}
	
	, VALIDATORS          => [['Spreadsheet lock validator', \&LockValidator]]
	, ERROR_HANDLER       => undef # user registred sub
	, MESSAGE             => 
				{
				ERROR => '#ERROR'
				, NEED_UPDATE => '#NEED UPDATE'
				, VIRTUAL_CELL => '#VC'
				, ROW_PREFIX => 'R-' 
				}
				
	, DEPENDENT_STACK     => []
	, CELLS               => {}
	) ;
}

#-------------------------------------------------------------------------------

sub Reset
{

=head2 Reset($self, [\%cells_setup_data])

Resets the contents of the spreadsheet and, optionally, sets optional spreadsheet parameters and
the contents of the spreadsheet cells

I<Arguments>  

=over 2

=item * $self - blessed Spreadsheet::Perl object

=item * \%cells_setup_data - Hash reference - Cells setup data, optional

=back

I<Returns> - Nothing

=cut

my ($self, $setup, $cell_data) = @_ ;

if(defined $setup)
	{
	if('HASH' eq ref $setup)
		{
		%$self = (GetSpreadsheetDefaultData(), %$setup) ;
		}
	else
		{
		confess "Setup data must be a hash reference!" 
		}
	}
	
if(defined $cell_data)
	{
	confess "cell data must be a hash reference!" unless 'HASH' eq ref $cell_data ;
	$self->{CELLS} = $cell_data ;
	}
else
	{
	$self->{CELLS} = {} ;
	}

return ;
}

#-------------------------------------------------------------------------------

sub TIEHASH 
{

=head2 [p] TIEHASH()

I<Arguments> - Automatically set by Perl

I<Returns> 

=over 2

=item * blessed Spreadsheet::Perl object

=back

=cut

my $class = shift ;

return($class) unless '' eq ref $class ;

my $self = 
	{
	  GetSpreadsheetDefaultData()
	, @_ 
	} ;

return(bless $self, $class) ;
}

#-------------------------------------------------------------------------------

sub DELETE   
{

=head2 i[p] DELETE($self, $address)

I<Arguments>

=over 2

=item * $self - blessed Spreadsheet::Perl object

=item * $address - cell address or cell address range

=back

I<Returns> - Nothing

=cut

my ($self, $address) = @_ ;

for my $current_address ($self->GetAddressList($address))
	{
	my $delete_cell = 1 ;

	if(exists $self->{CELLS}{$current_address}{DELETE_SUB})
		{
		$delete_cell = $self->{CELLS}{$current_address}{DELETE_SUB}->($self, $current_address, @{$self->{CELLS}{$current_address}{DELETE_SUB_ARGS}}) ;
		}

  if($delete_cell)
	  {
	  $self->MarkDependentForUpdate($self->{CELLS}{$current_address}, $current_address) ;

	  $self->InvalidateCellInDependent($self->GetName() . '!' . $current_address) ;

	  delete $self->{CELLS}{$current_address} ;
	  }
	}

return ;
}

#-------------------------------------------------------------------------------

sub CLEAR 
{

=head2 [p] CLEAR($self, $address)

I<Arguments>

=over 2

=item * $self - blessed Spreadsheet::Perl object

=item * $address - cell address or cell address range

=back

I<Returns> - Nothing

=cut

my ($self, $address) = @_ ; 

delete $self->{CELLS} ; 
# Todo: must call all set functions! and delete? functions? !!!!

return ;
}

#-------------------------------------------------------------------------------

sub EXISTS   
{

=head2 [p] EXISTS($self, $address)

Check fot the existance of a cell or all the cells in a range

I<Arguments>

=over 2

=item * $self - blessed Spreadsheet::Perl object

=item * $address - cell address or cell address range

=back

I<Returns> 0 if cell does not exists, 1 if cell exists

=cut

my ($self, $address) = @_ ; 

for my $current_address ($self->GetAddressList($address))
	{
	unless(exists $self->{CELLS}{$current_address})
		{
		return(0) ;
		}
	}
	
return(1) ;
}

#-------------------------------------------------------------------------------

sub FIRSTKEY 
{

=head2 [p] FIRSTKEY($self)

I<Arguments>

=over 2

=item * $self - blessed Spreadsheet::Perl object

=back

I<Returns> - See L<perltie> 

=cut

my $self = shift ;
scalar(keys %{$self->{CELLS}}) ;

return scalar each %{$self->{CELLS}} ;
}

#-------------------------------------------------------------------------------

sub NEXTKEY  
{

=head2 [p] NEXTKEY($self)

I<Arguments>

=over 2

=item * $self - blessed Spreadsheet::Perl object

=back

I<Returns> - See L<perltie> 

=cut

my $self = shift;
return scalar each %{ $self->{CELLS} }
}

#-------------------------------------------------------------------------------

sub DESTROY  
{

=head2 [p] DESTROY($self)

Not implemented.

=cut

}

#-------------------------------------------------------------------------------

sub LockValidator
{

=head2 GetInformation()

I<Arguments>

=over 2

=item * $self - blessed Spreadsheet::Perl object

=item * $address - cell address or cell address range

=back

I<Returns> - the spreadsheet or the cells lock status. Note: 0 if locked, 1 if not locked 

I<Warnings> - Carps if the lock is active

=cut


my ($self, $address) = @_ ;

if($self->{LOCKED})
	{
	carp "While setting '$address': Spreadsheet lock is active" ;
	return(0) ;
	}
else
	{
	if($self->IsCellLocked($address))
		{
		carp "While setting '$address': Cell lock is active" ;
			
		return(0) ;
		}
	else
		{
		return(1) ;
		}
	}

}

#-------------------------------------------------------------------------------

1 ;

