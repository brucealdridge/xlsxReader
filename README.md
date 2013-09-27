xlsxReader
==========

Simple php class to read in an xlsx file and parse it

First load in an xlsx file
```php
<?php
require 'xlsxReader.php';
$xlsxReader = new xlsxReader;
$xlsxReader->load('myfile.xlsx');
```

Then ...  

###Generate a csv file
```php
$csvfile = $xlsxReader->convert();
```

###Parse a xlsx file
```php
$xlsxReader->convert(function ($columns, $rowIndex) {
  echo $rowIndex.': '.explode(',', $columns).'<br>';
});
```
OR
```php
$xlsxReader->convert(function ($columns, $rowIndex) {
  echo $rowIndex.': '.explode(',', $columns).'<br>';
});

function rowParser($columns) {
    $db->query("INSERT INTO table VALUES ({$columns[1]}, {$columns[2]}, {$columns[3]})");
}
```
