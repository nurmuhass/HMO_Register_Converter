<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER["REQUEST_METHOD"] === "POST") {
    // Get data from the user input
    $data = $_POST["data"];
    $surnameData = $_POST["surnameData"];
    $firstNameData = $_POST["firstNameData"];
    $dobData = $_POST["dobData"];
    $genderData = $_POST["genderData"];
    $hmoData = $_POST["hmoData"];
    $hmoAcronymData = $_POST["hmoAcronymData"];

    // Separate the data into rows
    $dataRows = explode(PHP_EOL, $data);
    $surnameRows = explode(PHP_EOL, $surnameData);
    $firstNameRows = explode(PHP_EOL, $firstNameData);
    $dobRows = explode(PHP_EOL, $dobData);
    $genderRows = explode(PHP_EOL, $genderData);
    $hmoRows = explode(PHP_EOL, $hmoData);
    $hmoAcronymRows = explode(PHP_EOL, $hmoAcronymData);

    // Get HMO and HMO Acronym from user input
    $hmo = trim($hmoData);
    $hmoAcronym = trim($hmoAcronymData);

    // Initialize variables
    $outputData = [];

    // Process the data rows
    $currentNumber = null;
    foreach ($dataRows as $rowIndex => $rowData) {
        if (ctype_digit($rowData)) {
            $currentNumber = $rowData;
        } elseif (!empty(trim($rowData))) {
            $outputData[] = [
                $rowData,
                $currentNumber,
                $surnameRows[$rowIndex] ?? '',
                $firstNameRows[$rowIndex] ?? '',
                $dobRows[$rowIndex] ?? '',
                $genderRows[$rowIndex] ?? '',
                $hmoRows[$rowIndex] ?? $hmo,
                $hmoAcronymRows[$rowIndex] ?? $hmoAcronym,
                $rowData // Dependent_Type
            ];
        }
    }

    // Create a new spreadsheet
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Set headers
    $sheet->setCellValue('A1', 'enrollee_type');
    $sheet->setCellValue('B1', 'nhis_no');
    $sheet->setCellValue('C1', 'Surname');
    $sheet->setCellValue('D1', 'First Name');
    $sheet->setCellValue('E1', 'Date Of Birth');
    $sheet->setCellValue('F1', 'Sex');
    $sheet->setCellValue('G1', 'HMO');
    $sheet->setCellValue('H1', 'HMO Acronym');
    $sheet->setCellValue('I1', 'Dependent_Type'); // Add header for Dependent_Type
      $sheet->setCellValue('J1', 'Primary_Hcp');

    // Write the processed data to the spreadsheet
    foreach ($outputData as $rowNumber => $rowData) {
        foreach ($rowData as $columnNumber => $cellData) {
            $sheet->setCellValueByColumnAndRow($columnNumber + 1, $rowNumber + 2, $cellData);
        }
    }
  
    // Set headers to download the file as an Excel file
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="Converted_Register.xlsx"');
    header('Cache-Control: max-age=0');
    // Save the spreadsheet to a file or output to the browser
    $writer = new Xlsx($spreadsheet);
    $writer->save('php://output');
     header('Location:primaryhcp.php');
    exit;
  
}
?>


<!DOCTYPE html>
<html>

<head>
    <title>HMO Register Converter</title>
    <link rel="icon" href="logo.png" type="logo.png">
</head>

<body >

 <center> <h1>HMO Register Converter</h1></center>  
    <form method="post">
    
       <br> <h4 style="margin:0;align-items:  left">Enter HMO </h4>
         <input name="hmoData" rows="10" cols="50" placeholder="Enter HMO data" style="">
        <br>
        <h4 style="margin:1">Enter HMO Acronym</h4>
<input name="hmoAcronymData" rows="10" cols="50" placeholder="Enter HMO Acronym data" style="width:70%;height:100px">
        <br>
          <br>
           <h4 style="margin-bottom:1">Paste enrolle type and nhis no</h4>
        <textarea name="data" rows="10" cols="50" placeholder="Paste your data here" style=""></textarea><br>
        <h4>Paste Surname Column</h4>
        <textarea name="surnameData" rows="10" cols="50" placeholder="Paste surname data here" style="width:70%;height:300px"></textarea>
        <br>
        <h4>Paste FirstName Column</h4>
        <textarea name="firstNameData" rows="10" cols="50" placeholder="Paste firstNameData data here" style="width:70%;height:300px"></textarea>
        <br>
          <h4>Paste DOB Column</h4>
           <textarea name="dobData" rows="10" cols="50" placeholder="Paste dobData data here" style="width:70%;height:300px"></textarea>
        <br>
         <h4>Paste Gender Column</h4>
          <textarea name="genderData" rows="10" cols="50" placeholder="Paste gender data here" style="width:70%;height:300px"></textarea>

<br>
       
        <input type="submit" value="Process and Download">
    
        
        <button style="padding:7px;border-radius:5;margin-left:15%;"> <a href="primaryhcp.php" style="text-decoration: none;"> Process Primary HCP </a></button>
        
        
    </form>

</body>

</html>
<style>
    body{
        background-color:beige
    }
    textarea{
        width:70%;
        height:300px;
        border-radius: 5px;
        border-color: black;
        border-width:medium;
        text-align: center
    }
    input{
        width:70%;
        height:50px; 
        border-radius:5px;
    
   
    }
</style>
