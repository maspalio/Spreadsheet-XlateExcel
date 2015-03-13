on 'test' => sub {
  requires 'Test::Differences',       '0.5';
  requires 'Test::More',              '0.94';
};

on 'runtime' => sub {
  requires 'Carp::Assert::More',      '1.12';
  requires 'List::MoreUtils',         '0.406';
  requires 'Spreadsheet::ParseExcel', '0.58';
};
