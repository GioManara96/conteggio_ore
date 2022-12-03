<?php
require "vendor/autoload.php";
require 'Console/Table.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$inputFileName = "ore.xlsx";

if (file_exists($inputFileName)) {
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $spreadsheet = $reader->load("ore.xlsx");

    $tbl = createTable($spreadsheet);
    echo $tbl->getTable();

    echo "\nDesideri aggiungere una nuova giornata? (si/no)\n";
    $scelta = readline("Risposta: ");

    if ($scelta === "no") {
        echo "OK, buona giornata!\n";
    } else  if ($scelta === "si") {
        $giorno = readline("Inserisci data (dd/mm/yyyy): ");
        $ore = readline("Inserisci ore di lavoro: ");
        $confirm = readline("Confermi $giorno - $ore ore ? (si/no)");
        if ($confirm === "si") {
            $insert = [$giorno, $ore];
            insertNewRecord($spreadsheet, $insert);
            echo "\nGiornata registrata, continua così!\n";
        }
    } else {
        echo "Inserita risposta non corretta\n";
    }
} else {
    echo "File non presente. Un secondo, lo sto creando.\n";
    $spreadsheet = new Spreadsheet;
    $sheet = $spreadsheet->getActiveSheet();
    $headers = ["DATA", "ORE"];
    $sheet->fromArray($headers, NULL, "A1");

    $writer = new Xlsx($spreadsheet);
    $writer->save('ore.xlsx');
    echo "Fatto. Ora sei pronto a partire. Ri-esegui il programma per cominciare.\n";

}


/**
 * funzioni
 */

function createTable($spreadsheet) {
    $tbl = new Console_Table();
    $rows = $spreadsheet->getSheet(0)->toArray();

    // headers
    $headers = array();
    // dati
    $dati = array();
    for ($i = 0; $i < count($rows); $i++) {
        if ($i == 0) {
            foreach ($rows[0] as $header)
                array_push($headers, $header);
        } else {
            $tmp = array();
            foreach ($rows[$i] as $data) {
                array_push($tmp, $data);
            }
            array_push($dati, $tmp);
        }
    }

    $tbl->setHeaders($headers);
    // totale ore
    $tot_ore = 0;
    foreach ($dati as $giornata) {
        $tbl->addRow($giornata);
        $tot_ore += $giornata[1];
    }

    $tbl->addRow(array("", ""));
    $tbl->addRow(array("TOTALE", $tot_ore));

    return $tbl;
}

function insertNewRecord($spreadsheet, $data_to_insert) {
    $sheet = $spreadsheet->getActiveSheet();
    $rows = $sheet->getHighestRow() + 1;
    $sheet->insertNewRowBefore($rows);
    $sheet->fromArray($data_to_insert, NULL, "A$rows");

    $writer = new Xlsx($spreadsheet);
    $writer->save('ore.xlsx');
}
