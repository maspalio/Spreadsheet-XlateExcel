#
# Module.
#

package Spreadsheet::XlateExcel;

#
# Dependencies.
#

use Carp::Assert::More;
use Spreadsheet::ParseExcel;

#
# Bitch.
#

use warnings;
use strict;

#
# Documentation.
#

=head1 NAME

Spreadsheet::XlateExcel - Trigger a callback subroutine on each row of an Excel spreadsheet

=head1 VERSION

Version 0.01

=cut

our $VERSION = '0.01';

=head1 SYNOPSIS

This modules triggers a callback subroutine on each row of an Excel spreadsheet.

Wrote this module because I was fed up from writing the same boilerplate code ever when I had to mine spreadsheets for data.

Operates on every sheet unless a given sheet is targeted by name, RE inclusion or RE exclusion.

For example:

    use Spreadsheet::XlateExcel;

    my $id = Spreadsheet::XlateExcel->new ({ file => 'sheet.xls' });
    
    # rip odd rows of "Sheet2" sheet
    
    my $lol;
    
    $id->xlate ({
      on_sheet_named  => 'Sheet2',
      for_each_row_do => sub {
        my ( $sheet_id, $row, $row_vs ) = @_;
        
        push @$lol, $row_vs unless $row % 2;
      },
    });

=head1 METHODS

=cut

#
# Methods.
#

=head2 new

  my $id = Spreadsheet::XlateExcel->new ({ file => 'sheet.xls' })

Ye constructor.

=cut

sub new {
  my ( $class, $option ) = @_;
  
  assert_exists      $option=>'file';
  assert_nonblank    $option->{file};
  assert_defined  -f $option->{file}, 'incoming file exists';
  
  bless { book_id => Spreadsheet::ParseExcel->new->parse ( $option->{file} ) }, $class;
}

=head2 xlate

  $self->xlate ({ for_each_row_do => sub { my ( $sheet_id, $row, $row_vs ) = @_ ; ... } })

Applies for_each_row_do sub to each row of each sheet (unless filtered, see below) of the book.

Use on_sheet_named option to target a given book sheet by name.

Use on_sheets_like option to target a given book sheet by RE inclusion on name.

Use on_sheets_unlike option to target a given book sheet by RE exclusion on name.

Function gets called for each row, fed with L<Spreadsheet::ParseExcel::Worksheet> ID, row index and arrayref to row values parameters.

=cut

sub xlate {
  my ( $self, $option ) = @_;
  
  assert_exists $option => 'for_each_row_do';
    
  for my $sheet ( $self->book_id->worksheets ) {
    my $sheet_name = $sheet->get_name;
    
    next if $option->{on_sheet_named}   && $sheet_name ne $option->{on_sheet_named};
    next if $option->{on_sheets_like}   && $sheet_name !~ $option->{on_sheets_like};
    next if $option->{on_sheets_unlike} && $sheet_name =~ $option->{on_sheets_unlike};
    
    my ( $row_min, $row_max ) = $sheet->row_range;
    my ( $col_min, $col_max ) = $sheet->col_range;
    
    for my $row ( $row_min .. $row_max ) {      
      $option->{for_each_row_do}->( $sheet, $row, [ map { $sheet->get_cell( $row, $_ )->value } $col_min .. $col_max ] );
    }
  }
}

=head2 book_id

  my $book_id = $self->book_id ()

Accessor to L<Spreadsheet::ParseExcel::Workbook> instance ID.

=cut

sub book_id {
  my ( $self ) = @_;
  
  $self->{book_id};
}

#
# Documentation.
#

=head1 AUTHOR

Xavier Caron, C<< <xav at cpan.org> >>

=head1 BUGS

Please report any bugs or feature requests to C<bug-spreadsheet-xlateexcel at rt.cpan.org>, or through
the web interface at L<http://rt.cpan.org/NoAuth/ReportBug.html?Queue=Spreadsheet-XlateExcel>.  I will be notified, and then you'll
automatically be notified of progress on your bug as I make changes.

=head1 SUPPORT

You can find documentation for this module with the perldoc command.

    perldoc Spreadsheet::XlateExcel

You can also look for information at:

=over 4

=item * RT: CPAN's request tracker

L<http://rt.cpan.org/NoAuth/Bugs.html?Dist=Spreadsheet-XlateExcel>

=item * AnnoCPAN: Annotated CPAN documentation

L<http://annocpan.org/dist/Spreadsheet-XlateExcel>

=item * CPAN Ratings

L<http://cpanratings.perl.org/d/Spreadsheet-XlateExcel>

=item * Search CPAN

L<http://search.cpan.org/dist/Spreadsheet-XlateExcel/>

=back

=head1 ACKNOWLEDGEMENTS

=head1 LICENSE AND COPYRIGHT

Copyright 2010 Xavier Caron.

This program is free software; you can redistribute it and/or modify it
under the terms of either: the GNU General Public License as published
by the Free Software Foundation; or the Artistic License.

See http://dev.perl.org/licenses/ for more information.

=cut

#
# True.
#

1;
