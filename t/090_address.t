# test

use strict ;
use warnings ;

use Test::Exception ;
use Test::Warn;
use Test::NoWarnings qw(had_no_warnings);

use Test::More 'no_plan';
use Test::Block qw($Plan);

use Spreadsheet::Perl ; 

{
local $Plan = {'offset address' => 10} ;

my $ss = tie my %ss, "Spreadsheet::Perl" ;

for
	(
	  ['A1', 1, 1, 'B2']
	, ['Z9', 1, 0, 'AA9']
	, ['Z9', 1, 1, 'AA10']
	, ['ZZ1', 1, 1, 'AAA2']
	, ['AAA1', -1, 0, 'ZZ1']
	, ['ABC5', 25, 3, 'ACB8']
	, ['Z1', -25, 0, 'A1']
	, ['Z1', -26, 0, undef]
	, ['AA2', -26, -1, 'A1']
	, ['AA2', -26, -2, undef]
	)
	{
	my ($address, $column_offset, $row_offset, $expected_result) = @$_ ;
	my $offset_cell = $ss->OffsetAddress($address, $column_offset, $row_offset) ;
	is($offset_cell, $expected_result,'offsetting cell') ;
	}
}

=comment

{
local $Plan = {'' => } ;

is(result, expected, 'message') ;

throws_ok
	{
	
	} qr//, '' ;

lives_ok
	{
	
	} '' ;

like(result, qr//, '') ;

warning_like
	{
	} qr//i, '';

warnings_like
	{
	}
	[
	qr//i,
	qr//i,
	] '';


is_deeply
	(
	generated,
	[],
	''
	) ;


use Directory::Scratch ;
my $temp = Directory::Scratch->new();
my $dir  = $temp->mkdir('foo/bar');
my $file = $temp->touch('foo/bar/baz', qw(This is a file with lots of lines));


use Directory::Scratch::Structured qw(create_structured_tree piggyback_directory_scratch) ; 
my %tree_structure =
	(
	file_0 => [] ,
	
	dir_1 =>
		{
		subdir_1 =>{},
		file_1 =>[],
		file_a => [],
		},
	) ;

my $temporary_directory = create_structured_tree(%tree_structure) ;
$base = $temporary_directory->base() ;

my $scratch = Directory::Scratch->new;
$scratch->create_structured_tree(%tree_structure) ;
}

=cut
