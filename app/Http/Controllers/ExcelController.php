<?php

namespace App\Http\Controllers;

use App\Services\ExcelTransformer;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Storage;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ExcelController extends Controller
{
    protected ExcelTransformer $transformer;

    public function __construct(ExcelTransformer $transformer)
    {
        $this->transformer = $transformer;
    }

    public function index(Request $request)
    {
        $fileInfo = session('file_info');
        $allRawRows = session('raw_rows', []);
        $allTransformedRows = session('transformed_rows', []);
        $anomalies = session('anomalies', []);
        $rulesConfig = config('excel_rules');

        $activeSheet = $request->query('sheet', 'PUPW');

        $rawRows = $allRawRows[$activeSheet] ?? [];
        $transformedRows = $allTransformedRows[$activeSheet] ?? [];

        return view('workbench', compact('fileInfo', 'rawRows', 'transformedRows', 'anomalies', 'rulesConfig', 'activeSheet'));
    }

    public function upload(Request $request)
    {
        $request->validate([
            'excel_file' => 'required|file|mimes:xlsx,xls,csv,txt'
        ]);

        $file = $request->file('excel_file');

        // ============================================================
        // SIMPAN FAIL KE DIREKTORI SEMENTARA (TEMP) DALAM STOR LOKAL
        // ============================================================
        $path = $file->storeAs('temp', 'uploaded_' . time() . '_' . $file->getClientOriginalName());
        $realPath = Storage::path($path);

        try {
            $reader = IOFactory::createReaderForFile($realPath);
            $reader->setReadDataOnly(true);

            $spreadsheet = $reader->load($realPath);

            $sheetNames = ['PUPW', 'UTM-IDP', 'Foundation'];
            $allRawRows = [];
            $allTransformedRows = [];
            $allAnomalies = [];
            $totalRows = 0;
            $columns = [];

            foreach ($sheetNames as $sheetName) {
                $worksheet = $spreadsheet->getSheetByName($sheetName);
                if (!$worksheet) {
                    continue;
                }

                $highestRow = $worksheet->getHighestRow();
                $totalRows += ($highestRow - 1);

                // ============================================================
                // formatData=false DIGUNAKAN SUPAYA NOMBOR TELEFON TIDAK DITUKAR
                // KEPADA FORMAT SAINTIFIK SEMASA MEMBACA DATA DARI EXCEL
                // ============================================================
                $data = $worksheet->toArray(null, true, false, true);
                if (empty($data)) {
                    continue;
                }

                // ============================================================
                // BARIS PERTAMA ADALAH (HEADER)
                // ============================================================
                $headers = array_shift($data);
                $headers = array_map(function ($header) {
                    return strtoupper(trim((string)$header));
                }, $headers);

                if (empty($columns)) {
                    $columns = array_values(array_filter($headers));
                }

                // ============================================================
                // HADKAN KEPADA 20 BARIS PERTAMA SAHAJA UNTUK PREVIEW
                // ============================================================
                $previewData = array_slice($data, 0, 20);

                // ============================================================
                // MAPPING SETIAP BARIS DATA DENGAN NAMA (HEADER)
                // ============================================================
                $rawRows = [];
                foreach ($previewData as $rowIndex => $row) {
                    $mappedRow = [];
                    $hasValues = false;
                    foreach ($headers as $colKey => $colName) {
                        if (empty($colName)) continue;
                        $val = $row[$colKey] ?? '';
                        $mappedRow[$colName] = $val;
                        if ($val !== '') {
                            $hasValues = true;
                        }
                    }
                    if ($hasValues) {
                        $rawRows[$rowIndex + 2] = $mappedRow;
                    }
                }

                // ============================================================
                // JALANKAN TRANSFORMASI DATA UNTUK PREVIEW MENGGUNAKAN
                // transformSheet
                // ============================================================
                $transformedRows = $this->transformer->transformSheet($rawRows, $sheetName);

                // ============================================================
                // SAHKAN DATA DAN CARI SEBARANG ANOMALI / RALAT
                // ============================================================
                foreach ($rawRows as $rowNum => $row) {
                    $transformed = $transformedRows[$rowNum] ?? [];
                    $rowAnomalies = $this->validateRow($row, $transformed, $rowNum, $sheetName);
                    if (!empty($rowAnomalies)) {
                        foreach ($rowAnomalies as &$anomaly) {
                            $anomaly['field'] = '[' . $sheetName . '] ' . $anomaly['field'];
                        }
                        $allAnomalies = array_merge($allAnomalies, $rowAnomalies);
                    }
                }

                $allRawRows[$sheetName] = self::sanitizeArray($rawRows);
                $allTransformedRows[$sheetName] = $transformedRows;
            }

            if (empty($allRawRows)) {
                return back()->with('error', 'None of the required sheets (PUPW, UTM-IDP, Foundation) were found or they are empty.');
            }

            $fileInfo = [
                'name'       => $file->getClientOriginalName(),
                'path'       => $path,
                'total_rows' => $totalRows,
                'columns'    => $columns,
                'timestamp'  => now()->toDateTimeString(),
                'status'     => 'Parsed & Ready',
            ];

            session([
                'file_info'        => $fileInfo,
                'raw_rows'         => $allRawRows,
                'transformed_rows' => $allTransformedRows,
                'anomalies'        => $allAnomalies,
            ]);

            return redirect()->route('workbench.index')->with('success', 'File parsed successfully');
        } catch (\Exception $e) {
            return back()->with('error', 'Error reading Excel file: ' . $e->getMessage());
        }
    }

    public function dryRun(Request $request)
    {
        $fileInfo = session('file_info');
        if (!$fileInfo) {
            return back()->with('error', 'No file uploaded');
        }

        // ============================================================
        // TOGGLE STRICT MOD ATAU JALANKAN TRANSFORMASI MAYA (SIMULASI)
        // ============================================================
        $strict = $request->has('strict_mode') ? 'Strict' : 'Standard';

        return redirect()->route('workbench.index')->with('success', "Dry run complete (Mode: {$strict}). No database modifications written.");
    }

    public function reloadRules()
    {
        $fileInfo = session('file_info');
        if (!$fileInfo) {
            return back()->with('error', 'No file uploaded to apply rules to');
        }

        $allRawRows = session('raw_rows', []);
        $allTransformedRows = [];
        $allAnomalies = [];

        foreach ($allRawRows as $sheetName => $rawRows) {
            $transformedRows = $this->transformer->transformSheet($rawRows, $sheetName);
            $allTransformedRows[$sheetName] = $transformedRows;

            foreach ($rawRows as $rowNum => $row) {
                $transformed = $transformedRows[$rowNum] ?? [];
                $rowAnomalies = $this->validateRow($row, $transformed, $rowNum, $sheetName);
                if (!empty($rowAnomalies)) {
                    foreach ($rowAnomalies as &$anomaly) {
                        $anomaly['field'] = '[' . $sheetName . '] ' . $anomaly['field'];
                    }
                    $allAnomalies = array_merge($allAnomalies, $rowAnomalies);
                }
            }
        }

        session([
            'transformed_rows' => $allTransformedRows,
            'anomalies'        => $allAnomalies,
        ]);

        return redirect()->route('workbench.index')->with('success', 'Rules reloaded and applied to preview.');
    }

    public function export(Request $request)
    {
        $fileInfo = session('file_info');
        if (!$fileInfo) {
            return back()->with('error', 'No active session file');
        }

        $realPath = Storage::path($fileInfo['path']);

        try {
            $reader = IOFactory::createReaderForFile($realPath);
            $spreadsheet = $reader->load($realPath);

            $sheetNames = ['PUPW', 'UTM-IDP', 'Foundation'];

            foreach ($sheetNames as $sheetName) {
                $worksheet = $spreadsheet->getSheetByName($sheetName);
                if (!$worksheet) {
                    continue;
                }

                // ============================================================
                // formatData=false DIGUNAKAN SUPAYA NOMBOR TELEFON TIDAK DITUKAR
                // KEPADA FORMAT SAINTIFIK SEMASA MEMBACA DATA DARI EXCEL
                // ============================================================
                $data = $worksheet->toArray(null, true, false, true);
                $headers = array_shift($data);
                $headers = array_map(function ($header) {
                    return strtoupper(trim((string)$header));
                }, $headers);

                $cleanHeaders = [];
                foreach ($headers as $colKey => $colName) {
                    if (!empty($colName)) {
                        $cleanHeaders[$colKey] = $colName;
                    }
                }

                // ============================================================
                // CARI ATAU TAMBAH COLUMN ASA_KODTARAF JIKA TIDAK WUJUD
                // ============================================================
                $kodTarafColKey = array_search('ASA_KODTARAF', $headers);
                if ($kodTarafColKey === false) {
                    $highestCol = $worksheet->getHighestColumn();
                    $newColKey = $this->nextColumn($highestCol);
                    $worksheet->setCellValue($newColKey . '1', 'ASA_KODTARAF');
                    $kodTarafColKey = $newColKey;
                }

                // ============================================================
                // PROSES SETIAP BARIS DATA BERMULA DARI BARIS KE-2
                // (BARIS 1 ADALAH HEADER)
                // ============================================================
                foreach ($data as $index => $rowValues) {
                    $rowNum = $index + 2;

                    $assocRow = [];
                    foreach ($cleanHeaders as $colKey => $colName) {
                        $assocRow[$colName] = $rowValues[$colKey] ?? '';
                    }

                    $transformed = $this->transformer->transform($assocRow, $sheetName);

                    // ============================================================
                    // TULIS SEMUA COLUMN YANG TELAH DITRANSFORMASI KE DALAM SHEETS
                    // NOMBOR TELEFON MESTI DITULIS SEBAGAI (TYPE_STRING)
                    // UNTUK MENGELAKKAN FORMAT SAINTIFIK DALAM EXCEL
                    // ============================================================
                    $phoneFields = ['ASA_TELEFON', 'ASA_NOHP'];
                    foreach ($cleanHeaders as $colKey => $colName) {
                        if (array_key_exists($colName, $transformed)) {
                            $cellRef = $colKey . $rowNum;
                            if (in_array($colName, $phoneFields, true)) {
                                $worksheet->getCell($cellRef)
                                    ->setValueExplicit((string)$transformed[$colName], DataType::TYPE_STRING);
                            } else {
                                $worksheet->setCellValue($cellRef, $transformed[$colName]);
                            }
                        }
                    }

                    // ============================================================
                    // TULIS ASA_KODTARAF — MUNGKIN COLUMN BARU YANG BARU DITAMBAH
                    // ============================================================
                    $worksheet->setCellValue($kodTarafColKey . $rowNum, $transformed['ASA_KODTARAF'] ?? '');
                }
            }

            $writer = new Xlsx($spreadsheet);

            // ============================================================
            // TETAPKAN NAMA FAIL EKSPORT DI SINI
            // FORMAT: transformed_{TIMESTAMP}_{NAMA_FAIL_ASAL}
            // ============================================================
            // $exportName = 'transformed_' . time() . '_' . $fileInfo['name'];
            $exportName = '(CLEANUP) ' . $fileInfo['name'];
            $exportPath = 'temp/' . $exportName;

            Storage::makeDirectory('temp');
            $writer->save(Storage::path($exportPath));

            return Storage::download($exportPath, $exportName);
        } catch (\Exception $e) {
            return back()->with('error', 'Transformation export failed: ' . $e->getMessage());
        }
    }

    public function reset()
    {
        $fileInfo = session('file_info');
        if ($fileInfo && isset($fileInfo['path'])) {
            Storage::delete($fileInfo['path']);
        }

        session()->forget(['file_info', 'raw_rows', 'transformed_rows', 'anomalies']);
        return redirect()->route('workbench.index')->with('success', 'Session reset');
    }

    private function validateRow(array $raw, array $transformed, int $rowNum, string $sheetName): array
    {
        $anomalies = [];

        // ============================================================
        // SEMAK 1: PEMETAAN ASA_KAUM → ASA_KODTARAF
        // HANYA KAUM 'A' YANG PATUT MENGHASILKAN ASA_KODTARAF = 'B'
        // ============================================================
        $kaum = strtoupper(trim((string)($raw['ASA_KAUM'] ?? '')));
        if ($kaum !== '' && $kaum !== 'A' && !empty($transformed['ASA_KODTARAF'])) {
            $anomalies[] = [
                'row'         => $rowNum,
                'field'       => 'ASA_KAUM',
                'level'       => 'warning',
                'description' => "Kaum is '{$kaum}' (not 'A'). ASA_KODTARAF must be empty.",
            ];
        }

        // ============================================================
        // SEMAK 2: KOD SUBJEK ASA_KURTAWAR MESTI WUJUD DALAM MAPPING
        // KOD YANG TIDAK DIKENALI AKAN DITANDAKAN SEBAGAI RALAT
        // HANYA UNTUK SHEET PUPW DAN Foundation — UTM-IDP DIKECUALIKAN
        // ============================================================
        if ($sheetName !== 'UTM-IDP') {
            $kurTawar = trim((string)($raw['ASA_KURTAWAR'] ?? ''));
            if ($kurTawar !== '') {
                $mapping = config('excel_rules.kurtawar_mapping');
                if (!isset($mapping[$kurTawar])) {
                    $anomalies[] = [
                        'row'         => $rowNum,
                        'field'       => 'ASA_KURTAWAR',
                        'level'       => 'danger',
                        'description' => "Unmapped legacy subject code '{$kurTawar}'",
                    ];
                }
            }
        }

        return $anomalies;
    }

    private function nextColumn(string $column): string
    {
        $length = strlen($column);
        $lastChar = substr($column, -1);
        if ($lastChar === 'Z') {
            if ($length === 1) {
                return 'AA';
            } else {
                return $this->nextColumn(substr($column, 0, -1)) . 'A';
            }
        }
        return substr($column, 0, -1) . chr(ord($lastChar) + 1);
    }

    // ============================================================
    // SANITIZE NAN/INF FROM ANY NESTED ARRAY BEFORE SESSION/JSON
    // LAST-RESORT GUARD — HANDLES PHPSPREADSHEET FORMULA ERRORS
    // ============================================================
    private static function sanitizeArray(array $data): array
    {
        array_walk_recursive($data, function (&$val) {
            if (is_float($val) && (is_nan($val) || is_infinite($val))) {
                $val = null;
            }
        });
        return $data;
    }
}
