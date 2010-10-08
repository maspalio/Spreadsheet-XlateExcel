use strict;
use warnings;
use ExtUtils::MakeMaker;

WriteMakefile(
    NAME                => 'Spreadsheet::XlateExcel',
    AUTHOR              => q{Xavier Caron <xav@cpan.org>},
    VERSION_FROM        => 'lib/Spreadsheet/XlateExcel.pm',
    ABSTRACT_FROM       => 'lib/Spreadsheet/XlateExcel.pm',
    ($ExtUtils::MakeMaker::VERSION >= 6.3002
      ? ('LICENSE'=> 'perl')
      : ()),
    PL_FILES            => {},
    PREREQ_PM => {
        'Carp::Assert::More'      => 1.12,
        'Spreadsheet::ParseExcel' => 0.58,
        'Test::Most'              => 0.23,
    },
    dist                => { COMPRESS => 'gzip -9f', SUFFIX => 'gz', },
    clean               => { FILES => 'Spreadsheet-XlateExcel-*' },
);
