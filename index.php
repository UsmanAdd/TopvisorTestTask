<?php

use Classes\SpreadSheetTools;

require __DIR__ . '/vendor/autoload.php';
require __DIR__ . '/classes/SpreadSheetTools.php';

if (php_sapi_name() != 'cli') {
    throw new Exception('This application must be run on the command line.');
}

$title = 'Topvisor Test Task';
$range = 'A1:A10';
$valueInputOption = 'RAW';
$value = [[1], [2], [3], [4], [5], [6], [7], [8], [9], [10]];

$spreadsheet = new SpreadSheetTools();
$spreadsheetId = $spreadsheet->create($title);
$spreadsheet->updateValues($spreadsheetId, $range, $valueInputOption, $value);