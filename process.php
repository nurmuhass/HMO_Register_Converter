<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER["REQUEST_METHOD"] === "POST") {
    // Get data from the user input
    $data = $_POST["data"];

    // Separate the data into rows
    $rows = explode(PHP_EOL, $data);

    // Initialize variables
    $currentNumber = null;
    $outputData = [];

    // Process the rows
    foreach ($rows as $row) {
        // Check if the row is a number (unique number)
        if (ctype_digit($row)) {
            $currentNumber = $row;
        } elseif (!empty(trim($row))) {
            // If not a number and the row is not empty, it's data for an individual
            $outputData[$currentNumber][] = $row;
        }
    }

    // Create a new spreadsheet
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Write the headers
    $sheet->setCellValueByColumnAndRow(1, 1, 'enrollee_type');
    $sheet->setCellValueByColumnAndRow(2, 1, 'nhis_no');

    // Write the processed data to the spreadsheet
    $rowNumber = 2;
    foreach ($outputData as $number => $individualData) {
        foreach ($individualData as $rowData) {
            $sheet->setCellValueByColumnAndRow(1, $rowNumber, $rowData);
            $sheet->setCellValueByColumnAndRow(2, $rowNumber, $number);
            $rowNumber++;
        }
        $rowNumber++; // Add an extra line break
    }

    // Set headers to download the file as an Excel file
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header
