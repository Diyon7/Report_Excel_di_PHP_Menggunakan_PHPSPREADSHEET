<?php
include "koneksi.php";
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'No');
$sheet->setCellValue('B1', 'input');
$sheet->setCellValue('C1', 'nama');
$sheet->setCellValue('D1', 'jk');
$sheet->setCellValue('E1', 'nisn');
$sheet->setCellValue('F1', 'nik');
$sheet->setCellValue('G1', 'tempatlahir');
$sheet->setCellValue('H1', 'tanggallahir');
$sheet->setCellValue('I1', 'nomorakta');
$sheet->setCellValue('J1', 'agama');
$sheet->setCellValue('K1', 'negara');
$sheet->setCellValue('L1', 'kebutuhankhusus');
$sheet->setCellValue('M1', 'alamat');
$sheet->setCellValue('N1', 'rt');
$sheet->setCellValue('O1', 'rw');
$sheet->setCellValue('P1', 'dusun');
$sheet->setCellValue('Q1', 'desa');
$sheet->setCellValue('R1', 'kecamatan');
$sheet->setCellValue('S1', 'kp');
$sheet->setCellValue('T1', 'lintang');
$sheet->setCellValue('U1', 'bujur');
$sheet->setCellValue('V1', 'temting');
$sheet->setCellValue('W1', 'modtrans');
$sheet->setCellValue('X1', 'kks');
$sheet->setCellValue('Y1', 'anakke');
$sheet->setCellValue('Z1', 'pekps');
$sheet->setCellValue('AA1', 'kps');

$query = mysqli_query($koneksi, "select * from formulir");
$i = 2;
$no = 1;

while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $no++);
    $sheet->setCellValue('B' . $i, $row['input']);
    $sheet->setCellValue('C' . $i, $row['nama']);
    $sheet->setCellValue('D' . $i, $row['jk']);
    $sheet->setCellValue('E' . $i, $row['nisn']);
    $sheet->setCellValue('F' . $i, $row['nik']);
    $sheet->setCellValue('G' . $i, $row['tempatlahir']);
    $sheet->setCellValue('H' . $i, $row['tanggallahir']);
    $sheet->setCellValue('I' . $i, $row['nomorakta']);
    $sheet->setCellValue('J' . $i, $row['agama']);
    $sheet->setCellValue('K' . $i, $row['negara']);
    $sheet->setCellValue('L' . $i, $row['kebutuhankhusus']);
    $sheet->setCellValue('M' . $i, $row['alamat']);
    $sheet->setCellValue('N' . $i, $row['rt']);
    $sheet->setCellValue('O' . $i, $row['rw']);
    $sheet->setCellValue('P' . $i, $row['dusun']);
    $sheet->setCellValue('Q' . $i, $row['desa']);
    $sheet->setCellValue('R' . $i, $row['kecamatan']);
    $sheet->setCellValue('S' . $i, $row['kp']);
    $sheet->setCellValue('T' . $i, $row['lintang']);
    $sheet->setCellValue('U' . $i, $row['bujur']);
    $sheet->setCellValue('V' . $i, $row['temting']);
    $sheet->setCellValue('W' . $i, $row['modtrans']);
    $sheet->setCellValue('X' . $i, $row['kks']);
    $sheet->setCellValue('Y' . $i, $row['anakke']);
    $sheet->setCellValue('Z' . $i, $row['pekps']);
    $sheet->setCellValue('AA' . $i, $row['kps']);
    $i++;
}

$styleArray = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ]
];
$i = $i - 1;
$sheet->getStyle('A1:AA' . $i)->applyFromArray($styleArray);

$writer = new Xlsx($spreadsheet);
$writer->save("Formulir.xlsx");
header("Location:../formulir_peserta_didik.php");
