<?php

namespace App\Services;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

class ExcelReaderService
{
    // ============================================================
    // BACA HELAIAN-HELAIAN YANG DITENTUKAN DARI FAIL EXCEL
    // KEMBALIKAN DATA DALAM BENTUK ARRAY BERSUSUN MENGIKUT NAMA HELAIAN
    // ============================================================
    public function read(string $filePath, array $sheetNames): array
    {
        $spreadsheet = IOFactory::load($filePath);
        $data        = [];

        foreach ($sheetNames as $targetName) {
            $sheet = null;

            // ============================================================
            // CARI HELAIAN MENGGUNAKAN PADANAN TIDAK SENSITIF KES (CASE-INSENSITIVE)
            // ============================================================
            foreach ($spreadsheet->getAllSheets() as $s) {
                if (strcasecmp($s->getTitle(), $targetName) === 0) {
                    $sheet = $s;
                    break;
                }
            }

            if (!$sheet) {
                continue;
            }

            $highestRow         = $sheet->getHighestRow();
            $highestColumn      = $sheet->getHighestColumn();
            $highestColumnIndex = Coordinate::columnIndexFromString($highestColumn);

            if ($highestRow < 1) {
                continue;
            }

            // ============================================================
            // BACA KEPALA LAJUR (HEADER) DARI BARIS PERTAMA
            // JIKA SEL MENGANDUNGI FORMULA, KIRA HASILNYA
            // ============================================================
            $headers = [];
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
                $headers[$col] = $val !== null ? trim((string)$val) : '';
            }

            // ============================================================
            // LANGKAU HELAIAN JIKA TIADA KEPALA LAJUR
            // ============================================================
            if (empty(array_filter($headers))) {
                continue;
            }

            // ============================================================
            // BACA SETIAP BARIS DATA BERMULA DARI BARIS KE-2
            // SIMPAN HANYA BARIS YANG MENGANDUNGI SEKURANG-KURANG SATU NILAI
            // ============================================================
            $rows = [];
            for ($row = 2; $row <= $highestRow; $row++) {
                $rowData = [];
                $hasData = false;
                for ($col = 1; $col <= $highestColumnIndex; $col++) {
                    $header = $headers[$col] ?? '';
                    if ($header === '') {
                        continue;
                    }
                    $cell = $sheet->getCell([$col, $row]);
                    $val  = $cell->getValue();
                    if (is_string($val) && str_starts_with($val, '=')) {
                        try {
                            $val = $cell->getCalculatedValue();
                        } catch (\Exception $e) {
                            $val = $cell->getOldCalculatedValue();
                        }
                    }
                    $rowData[$header] = $val;
                    if ($val !== null && $val !== '') {
                        $hasData = true;
                    }
                }

                // ============================================================
                // SIMPAN BARIS JIKA ADA DATA
                // ============================================================
                if ($hasData) {
                    $rows[$row] = $rowData;
                }
            }

            $data[$targetName] = [
                'headers' => array_values(array_filter($headers)),
                'rows'    => $rows,
            ];
        }

        return $data;
    }
}
