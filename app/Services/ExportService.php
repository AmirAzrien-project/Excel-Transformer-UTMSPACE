<?php

namespace App\Services;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use Illuminate\Support\Facades\Storage;

class ExportService
{
    // ============================================================
    // EKSPORT DATA YANG TELAH DITRANSFORMASI KE DALAM FAIL EXCEL
    // DATA DITULIS SEMULA KE DALAM HELAIAN ASAL FAIL YANG DIMUAT NAIK
    // FAIL BARU DISIMPAN DALAM DIREKTORI SEMENTARA (TEMP)
    // ============================================================
    public function export(string $sourceRealPath, array $processedSheets): string
    {
        $spreadsheet = IOFactory::load($sourceRealPath);

        foreach ($processedSheets as $sheetName => $rows) {
            $sheet = null;

            // ============================================================
            // CARI HELAIAN MENGGUNAKAN PADANAN TIDAK SENSITIF KES
            // ============================================================
            foreach ($spreadsheet->getAllSheets() as $s) {
                if (strcasecmp($s->getTitle(), $sheetName) === 0) {
                    $sheet = $s;
                    break;
                }
            }

            if (!$sheet) {
                continue;
            }

            $highestColumn      = $sheet->getHighestColumn();
            $highestColumnIndex = Coordinate::columnIndexFromString($highestColumn);

            // ============================================================
            // BINA PETA: NAMA KEPALA LAJUR → NOMBOR LAJUR
            // ============================================================
            $headerToCol = [];
            for ($col = 1; $col <= $highestColumnIndex; $col++) {
                $cell = $sheet->getCell([$col, 1]);
                $val  = $cell->getValue();
                if (is_string($val) && str_starts_with($val, '=')) {
                    try {
                        $val = $cell->getCalculatedValue();
                    } catch (\Exception $e) {
                        $val = $cell->getOldCalculatedValue();
                    }
                }
                $header = $val !== null ? trim((string)$val) : '';
                if ($header !== '') {
                    $headerToCol[$header] = $col;
                }
            }

            // ============================================================
            // TULIS DATA SETIAP BARIS SEMULA KE DALAM HELAIAN EXCEL
            // ============================================================
            foreach ($rows as $rowNum => $rowData) {
                foreach ($rowData as $header => $value) {
                    if (isset($headerToCol[$header])) {
                        $col = $headerToCol[$header];
                        $sheet->setCellValue([$col, $rowNum], $value);
                    }
                }
            }
        }

        // ============================================================
        // JANA NAMA FAIL EKSPORT SECARA DINAMIK
        // FORMAT: transformed_{TIMESTAMP}_{NAMA_FAIL_ASAL}
        // ============================================================
        $filename           = 'transformed_' . time() . '_' . basename($sourceRealPath);
        $relativeExportPath = 'temp/' . $filename;
        $absoluteExportPath = Storage::path($relativeExportPath);

        // ============================================================
        // PASTIKAN DIREKTORI WUJUD SEBELUM MENYIMPAN FAIL
        // ============================================================
        $dir = dirname($absoluteExportPath);
        if (!is_dir($dir)) {
            mkdir($dir, 0755, true);
        }

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($absoluteExportPath);

        return $relativeExportPath;
    }
}
