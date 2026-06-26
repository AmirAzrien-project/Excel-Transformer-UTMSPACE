<?php

namespace App\Services;

class ExcelTransformer
{
    protected array $rules;

    public function __construct()
    {
        $this->rules = config('excel_rules');
    }

    // ============================================================
    // TRANSFORMASI SEMUA BARIS DALAM SATU HELAIAN (SHEET)
    // ============================================================
    public function transformSheet(array $rows, string $sheetName): array
    {
        $transformed = [];
        foreach ($rows as $key => $row) {
            $transformed[$key] = $this->transform($row, $sheetName);
        }
        return $transformed;
    }

    // ============================================================
    // TRANSFORMASI SATU BARIS DATA — SEMUA PERATURAN DIJALANKAN
    // MENGIKUT SUSUNAN YANG DITETAPKAN
    // ============================================================
    public function transform(array $row, ?string $sheetName = null): array
    {
        if ($sheetName !== 'UTM-IDP') {
            $row = $this->transformKurTawar($row);
        }
        $row = $this->transformKodTaraf($row);
        $row = $this->normalizeDatesInRow($row);
        $row = $this->cleanNamesInRow($row);
        $row = $this->splitAddress($row);
        $row = $this->detectForeigner($row);
        $row = $this->cleanBandar($row);
        $row = $this->cleanPhone($row);
        $row = $this->cleanKodKolej($row);
        $row = $this->sanitizeForJson($row);

        return $row;
    }

    // ============================================================
    // PETA KAUM → KOD TARAF
    // KAUM A, D, E, F → ASA_KODTARAF = 'B'
    // SELAIN ITU → ASA_KODTARAF = ''
    // ============================================================
    private function transformKodTaraf(array $row): array
    {
        $kaum = strtoupper(trim((string)($row['ASA_KAUM'] ?? '')));

        if (in_array($kaum, ['A', 'D', 'E', 'F'], true)) {
            $row['ASA_KODTARAF'] = 'B';
        } else {
            $row['ASA_KODTARAF'] = '';
        }

        return $row;
    }

    // ============================================================
    // PETA KOD KURSUS LAMA (ASA_KURTAWAR) KEPADA KOD BAHARU
    // PEMETAAN DIAMBIL DARI config/excel_rules.php
    // ============================================================
    private function transformKurTawar(array $row): array
    {
        $kod = trim((string)($row['ASA_KURTAWAR'] ?? ''));

        if (isset($this->rules['kurtawar_mapping'][$kod])) {
            $row['ASA_KURTAWAR'] = $this->rules['kurtawar_mapping'][$kod];
        }

        return $row;
    }

    // ============================================================
    // NORMALISASI TARIKH UNTUK MEDAN ASA_TKHDAFTAR DAN ASA_TKHLAHIR
    // ============================================================
    private function normalizeDatesInRow(array $row): array
    {
        foreach (['ASA_TKHDAFTAR', 'ASA_TKHLAHIR'] as $field) {
            if (isset($row[$field])) {
                $row[$field] = $this->normalizeDate((string)$row[$field]);
            }
        }
        return $row;
    }

    // ============================================================
    // TUKAR PELBAGAI FORMAT TARIKH KEPADA FORMAT PIAWAI: YYYY-MM-DD
    // SOKONGAN: NOMBOR SIRI EXCEL, DD/MM/YYYY, DD-MM-YYYY, DLL.
    // ============================================================
    private function normalizeDate(string $val): string
    {
        $val = trim($val);
        if ($val === '') {
            return '';
        }

        // ============================================================
        // SEMAK NOMBOR SIRI EXCEL (CONTOH: 45000 = 2023-03-15)
        // ============================================================
        if (is_numeric($val) && $val > 0 && strpos($val, '.') === false) {
            try {
                $dt = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject((int)$val);
                return $dt->format('Y-m-d');
            } catch (\Exception $e) {
                $timestamp = ((int)$val - 25569) * 86400;
                return gmdate("Y-m-d", $timestamp);
            }
        }

        // ============================================================
        // CUBA PELBAGAI FORMAT TARIKH DALAM BENTUK TEKS
        // ============================================================
        $formats = ['d/m/Y', 'd-m-Y', 'Y-m-d', 'Y/m/d', 'j/n/Y', 'j-n-Y'];
        foreach ($formats as $format) {
            $dt = \DateTime::createFromFormat($format, $val);
            if ($dt !== false) {
                return $dt->format('Y-m-d');
            }
        }

        $ts = strtotime($val);
        if ($ts !== false && $ts > 0) {
            return date('Y-m-d', $ts);
        }

        return '';
    }

    // ============================================================
    // BERSIHKAN NAMA (ASA_NAMA) — BUANG SIMBOL, HURUF BESAR
    // ============================================================
    private function cleanNamesInRow(array $row): array
    {
        if (isset($row['ASA_NAMA'])) {
            $row['ASA_NAMA'] = $this->cleanName($row['ASA_NAMA']);
        }
        return $row;
    }

    // ============================================================
    // PROSES PEMBERSIHAN NAMA:
    // 1. GANTI @ DAN KOMA DENGAN RUANG (UNTUK PISAHKAN PERKATAAN)
    // 2. BUANG SEMUA SIMBOL LAIN
    // 3. RUNTUHKAN RUANG BERGANDA
    // 4. TUKAR KEPADA HURUF BESAR
    // ============================================================
    private function cleanName(string $name): string
    {
        $name = str_replace(['@', ','], ' ', $name);
        $cleaned = preg_replace('/[^A-Za-z\s]/', '', $name);
        $cleaned = preg_replace('/\s+/', ' ', $cleaned);
        return strtoupper(trim($cleaned));
    }

    // ============================================================
    // KESAN PEMOHON ASING BERDASARKAN:
    // 1. FORMAT IC (NRIC) — 12 DIGIT = MALAYSIA
    // 2. FORMAT NOMBOR TELEFON — 01x / 60x = MALAYSIA
    // 3. LOKASI ASING DALAM ALAMAT / BANDAR / NEGARA
    //
    // TINDAKAN: JIKA DIKENAL PASTI ASING DAN ASA_NEGARA = 'M'/'MAL'
    //           → KOSONGKAN ASA_NEGARA SAHAJA
    //           JANGAN UBAH MANA-MANA MEDAN LAIN
    // ============================================================
    private function detectForeigner(array $row): array
    {
        $nokp    = trim((string)($row['ASA_NOKP']    ?? ''));
        $phone   = (string)($row['ASA_TELEFON']      ?? '');
        $nohp    = (string)($row['ASA_NOHP']         ?? '');
        $alamat1 = (string)($row['ASA_ALAMAT1']      ?? '');
        $alamat2 = (string)($row['ASA_ALAMAT2']      ?? '');
        $bandar  = (string)($row['ASA_BANDAR']       ?? '');
        $negara  = (string)($row['ASA_NEGARA']       ?? '');

        // ============================================================
        // SEMAK 1: IC 12 DIGIT = MALAYSIA
        // ============================================================
        $cleanNokp    = preg_replace('/[^0-9]/', '', $nokp);
        $resemblesNric = (strlen($cleanNokp) === 12);

        // ============================================================
        // SEMAK 2: NOMBOR TELEFON BERMULA 01 ATAU 60 = MALAYSIA
        // ============================================================
        $resemblesPhone = false;
        foreach ([$phone, $nohp] as $rawPhone) {
            $p = preg_replace('/[^0-9]/', '', $rawPhone);
            if ($p !== '' && (str_starts_with($p, '01') || str_starts_with($p, '60'))) {
                $resemblesPhone = true;
            }
        }

        // ============================================================
        // SEMAK 3: CARI NAMA LOKASI ASING DALAM TEKS ALAMAT
        // ============================================================
        $addressText = strtoupper($alamat1 . ' ' . $alamat2 . ' ' . $bandar . ' ' . $negara);
        $foreignLocations = [
            'SINGAPORE', 'JAKARTA', 'INDONESIA', 'BANGKOK', 'THAILAND',
            'BEIJING', 'CHINA', 'INDIA', 'PAKISTAN', 'BANGLADESH'
        ];
        $containsForeignLocation = false;
        foreach ($foreignLocations as $loc) {
            if (str_contains($addressText, $loc)) {
                $containsForeignLocation = true;
                break;
            }
        }

        // ============================================================
        // KLASIFIKASI: BUKAN MALAYSIA JIKA —
        // (TIADA IC MALAYSIA DAN TIADA TELEFON MALAYSIA) ATAU ADA LOKASI ASING
        // ============================================================
        $isNonMalaysian = false;
        if ($containsForeignLocation) {
            $isNonMalaysian = true;
        } elseif (!$resemblesNric && !$resemblesPhone) {
            $isNonMalaysian = true;
        }

        // ============================================================
        // TINDAKAN: KOSONGKAN ASA_NEGARA JIKA BUKAN MALAYSIA
        // DAN ASA_NEGARA MENGANDUNGI TEPAT 'M' ATAU 'MAL'
        // ============================================================
        if ($isNonMalaysian) {
            $negaraUpper = strtoupper(trim($negara));
            if ($negaraUpper === 'M' || $negaraUpper === 'MAL') {
                $row['ASA_NEGARA'] = '';
            }
        }

        return $row;
    }

    // ============================================================
    // BERSIHKAN NOMBOR TELEFON (ASA_TELEFON DAN ASA_NOHP)
    // 1. TANGANI NILAI FLOAT (NOTASI SAINTIFIK DARI EXCEL)
    // 2. BUANG SEMUA BUKAN DIGIT (+, -, RUANG, KURUNGAN, DLL.)
    // 3. HADKAN KEPADA MAKSIMUM 13 DIGIT
    // 4. SIMPAN SEBAGAI STRING — JANGAN TUKAR KEPADA INT/FLOAT
    // ============================================================
    private function cleanPhone(array $row): array
    {
        foreach (['ASA_TELEFON', 'ASA_NOHP'] as $field) {
            if (isset($row[$field]) && $row[$field] !== null && $row[$field] !== '') {
                $val = $row[$field];

                // ============================================================
                // JIKA NILAI ADALAH FLOAT/INT (NOTASI SAINTIFIK DARI PHPSPREADSHEET),
                // TUKAR KEPADA STRING PENUH TANPA NOTASI SAINTIFIK
                // ============================================================
                if (is_float($val) || is_int($val)) {
                    $val = number_format($val, 0, '.', '');
                } else {
                    $val = (string)$val;
                }

                $digits        = preg_replace('/[^0-9]/', '', $val);
                $row[$field]   = substr($digits, 0, 13);
            } else {
                $row[$field] = '';
            }
        }
        return $row;
    }

    // ============================================================
    // BAHAGIKAN ALAMAT JIKA ASA_ALAMAT1 MELEBIHI PANJANG MAKSIMUM
    // LIMPAHAN AKAN DIPINDAHKAN KE ASA_ALAMAT2
    // PANJANG MAKSIMUM BOLEH DIKONFIGURASI (DEFAULT: 40 AKSARA)
    // ============================================================
    private function splitAddress(array $row): array
    {
        $maxLen  = config('excel_rules.alamat1_max_length', 40);
        $alamat1 = isset($row['ASA_ALAMAT1']) ? (string)$row['ASA_ALAMAT1'] : '';
        $alamat2 = isset($row['ASA_ALAMAT2']) ? (string)$row['ASA_ALAMAT2'] : '';

        // ============================================================
        // NORMALISASI PEMISAH BARIS
        // ============================================================
        $alamat1 = str_replace(["\r\n", "\r"], "\n", $alamat1);
        $alamat1 = trim($alamat1);

        $alamat2 = str_replace(["\r\n", "\r"], "\n", $alamat2);
        $alamat2 = trim($alamat2);

        if ($alamat1 === '') {
            return $row;
        }

        // ============================================================
        // BAHAGIKAN MENGIKUT BARIS DAHULU
        // ============================================================
        $lines          = explode("\n", $alamat1);
        $alamat1Lines   = [];
        $overflowLines  = [];
        $buildingAlamat1 = true;

        foreach ($lines as $line) {
            $line = trim($line);
            if ($line === '') {
                continue;
            }

            if ($buildingAlamat1) {
                // ============================================================
                // JIKA BARIS PERTAMA TERLALU PANJANG, BAHAGIKAN MENGIKUT PERKATAAN
                // (JANGAN POTONG DI TENGAH PERKATAAN)
                // ============================================================
                if (empty($alamat1Lines) && strlen($line) > $maxLen) {
                    $words        = explode(' ', $line);
                    $addedWords   = [];
                    $overflowWords = [];
                    $buildingWords = true;

                    foreach ($words as $word) {
                        $word = trim($word);
                        if ($word === '') {
                            continue;
                        }

                        if ($buildingWords) {
                            $testStr = empty($addedWords) ? $word : implode(' ', $addedWords) . ' ' . $word;
                            if (strlen($testStr) <= $maxLen) {
                                $addedWords[] = $word;
                            } else {
                                $overflowWords[] = $word;
                                $buildingWords   = false;
                            }
                        } else {
                            $overflowWords[] = $word;
                        }
                    }

                    $alamat1Lines[] = implode(' ', $addedWords);
                    if (!empty($overflowWords)) {
                        $overflowLines[] = implode(' ', $overflowWords);
                    }
                    $buildingAlamat1 = false;
                } else {
                    // ============================================================
                    // CUBA TAMBAH BARIS KE ASA_ALAMAT1 JIKA MASIH MUAT
                    // ============================================================
                    $testStr = empty($alamat1Lines) ? $line : implode("\n", $alamat1Lines) . "\n" . $line;
                    if (strlen($testStr) <= $maxLen) {
                        $alamat1Lines[] = $line;
                    } else {
                        $overflowLines[] = $line;
                        $buildingAlamat1 = false;
                    }
                }
            } else {
                $overflowLines[] = $line;
            }
        }

        $newAlamat1  = implode("\n", $alamat1Lines);
        $overflowText = implode("\n", $overflowLines);

        // ============================================================
        // BUANG KOMA DI HUJUNG / PERMULAAN SEMPADAN
        // ============================================================
        $newAlamat1  = rtrim(trim($newAlamat1), ',');
        $overflowText = ltrim(trim($overflowText), ',');

        $row['ASA_ALAMAT1'] = $newAlamat1;

        if ($overflowText !== '') {
            if ($alamat2 !== '') {
                $row['ASA_ALAMAT2'] = $alamat2 . "\n" . $overflowText;
            } else {
                $row['ASA_ALAMAT2'] = $overflowText;
            }
        } else {
            $row['ASA_ALAMAT2'] = $alamat2;
        }

        return $row;
    }

    // ============================================================
    // BERSIHKAN DAN NORMALISASI NAMA BANDAR (ASA_BANDAR)
    // ============================================================
    private function cleanBandar(array $row): array
    {
        if (isset($row['ASA_BANDAR'])) {
            $row['ASA_BANDAR'] = $this->normalizeCity($row['ASA_BANDAR']);
        }
        return $row;
    }

    // ============================================================
    // NORMALISASI NAMA BANDAR:
    // 1. TUKAR KE HURUF BESAR
    // 2. BUANG TEKS SELEPAS KOMA / SIMBOL PEMISAH
    // 3. GUNAKAN PEMETAAN TEPAT JIKA WUJUD
    // 4. GUNAKAN SINGKATAN GENERIK JIKA TIADA PEMETAAN TEPAT
    // ============================================================
    private function normalizeCity(string $bandar): string
    {
        $bandar = trim(strtoupper($bandar));
        if ($bandar === '') {
            return '';
        }

        // ============================================================
        // BUANG TEKS SELEPAS SIMBOL PEMISAH (KOMA, KURUNGAN, DLL.)
        // ============================================================
        $parts  = preg_split('/[,;\/|\\\(\[-]/u', $bandar);
        $bandar = trim($parts[0]);

        // ============================================================
        // PEMETAAN TEPAT — NAMA BANDAR PENUH → SINGKATAN RASMI
        // ============================================================
        $exactMappings = [
            'BANDAR PUTRA'                 => 'BANDAR PUTRA',
            'KUALA LUMPUR'                 => 'KUALA LUMPUR',
            'BANDAR BARU PERMAS JAYA'      => 'BB PERMAS JAYA',
            'PRESINT 8 PUTRAJAYA'          => 'PUTRAJAYA',
            'PRESINT 9 PUTRAJAYA'          => 'PUTRAJAYA',
            'PRESINT 11 PUTRAJAYA'         => 'PUTRAJAYA',
            'BUKIT MERTAJAM'               => 'BKT MERTAJAM',
            'TELOK PANGLIMA GARANG'        => 'TELOK PANGLIMA GRG',
            'KUALA TERENGGANU'             => 'K. TERENGGANU',
            'KUALA KUBU BARU'              => 'K. KUBU BARU',
            'BANDAR BARU NILAI'            => 'B. BARU NILAI',
            'BANDAR PUNCAK ALAM'           => 'BDR. PUNCAK ALAM',
            'WAKAF BHARU TUMPAT'           => 'WKF BARU TUMPAT',
            'BANDAR MAHKOTA CHERAS'        => 'BDR. MAHKOTA CHERAS',
            'KOTA KINABATANGAN'            => 'KT. KINABATANGAN',
            'BANDAR AL MUKTAFI BILLAH SHAH' => 'B. AL MUKTAFI BILLAH SHAH',
            'AYER BALOI PONTIAN'           => 'PONTIAN',
            'BANDAR BARU BANGI'            => 'BB BANGI',
            'TAMAN TAMPOI UTAMA'           => 'TMN TAMPOI UTAMA',
            'TANJONG RAMBUTAN'             => 'TJG RAMBUTAN',
            'BANDAR BARU ENSTEK'           => 'BB ENSTEK',
            'BANDAR PENAWAR'               => 'BDR PENAWAR',
            'BANDAR MELAKA TENGAH'         => 'MELAKA TENGAH',
            'BANDAR SERI JEMPOL'           => 'B SERI JEMPOL',
        ];

        if (isset($exactMappings[$bandar])) {
            return $exactMappings[$bandar];
        }

        // ============================================================
        // SINGKATAN GENERIK — GANTIKAN KATA KUNCI BIASA
        // ============================================================
        $genericReplacements = [
            '/\bBANDAR BARU\b/' => 'BB',
            '/\bBANDAR\b/'      => 'BDR',
            '/\bKUALA\b/'       => 'K.',
            '/\bKOTA\b/'        => 'KT.',
            '/\bTAMAN\b/'       => 'TMN',
            '/\bTANJONG\b/'     => 'TJG',
            '/\bTANJUNG\b/'     => 'TJG',
            '/\bWAKAF\b/'       => 'WKF',
            '/\bBUKIT\b/'       => 'BKT',
        ];

        foreach ($genericReplacements as $pattern => $replacement) {
            $bandar = preg_replace($pattern, $replacement, $bandar);
        }

        return trim(preg_replace('/\s+/', ' ', $bandar));
    }

    // ============================================================
    // KOSONGKAN ASA_KODKOLEJ — MEDAN INI TIDAK DIPERLUKAN
    // ============================================================
    private function cleanKodKolej(array $row): array
    {
        if (array_key_exists('ASA_KODKOLEJ', $row)) {
            $row['ASA_KODKOLEJ'] = '';
        }
        return $row;
    }

    // ============================================================
    // SANITIZE ARRAY FOR JSON — CONVERT NAN/INF FLOATS TO NULL
    // PHPSPREADSHEET CAN RETURN NAN FOR FORMULA ERRORS (#DIV/0! ETC)
    // PHP json_encode THROWS JSON_ERROR_INF_OR_NAN ON THESE VALUES
    // ============================================================
    private function sanitizeForJson(array $row): array
    {
        foreach ($row as $key => $value) {
            if (is_float($value) && (is_nan($value) || is_infinite($value))) {
                $row[$key] = null;
            } elseif (is_array($value)) {
                $row[$key] = $this->sanitizeForJson($value);
            }
        }
        return $row;
    }
}
