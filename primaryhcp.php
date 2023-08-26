<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER["REQUEST_METHOD"] == "POST") {
    $input = $_POST['input'];
    $lines = explode("\n", $input);

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle('Primary HCP Codes');

    // Add the header row
    $sheet->setCellValue('A1', 'Primary HCP');

    $row = 2; // Start from the second row after the header
    foreach ($lines as $line) {
        $data = explode('.', $line);
        if (count($data) == 2) {
            $code = trim($data[0]);
            $repeat = intval(trim($data[1]));

            for ($i = 0; $i < $repeat; $i++) {
                $sheet->setCellValue('A' . $row, $code);
                $row++;
            }
        }
    }

    $writer = new Xlsx($spreadsheet);
    $filename = 'primary_hcp_codes.xlsx';

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="' . $filename . '"');
    header('Cache-Control: max-age=0');

    $writer->save('php://output');
  
    exit;
}
?>

<!DOCTYPE html>
<html>
<head>

    <title>HMO Register Converter(Primary HCP)</title>
    <link rel="icon" href="logo.png" type="logo.png">
</head>
<body style=" background-color:beige">
    <h2>Enter Primary HCP Codes and Repetitions</h2>
    <form method="post">
        <textarea name="input" rows="10" cols="50" placeholder="Paste primary HCP codes and repetitions here" style="     width:70%;
        height:300px;
        border-radius: 5px;
        border-color: black;
        border-width: thick;
        text-align: center"></textarea><br>
        <center style="width:70%;" ><input type="submit" value="Generate Excel" style="width:70%;height:50"> </center> 
    </form>
</body>
</html>

