<?php

namespace Tests\Unit;

use Tests\TestCase;
use App\Services\ExcelTransformer;

class ExcelTransformerTest extends TestCase
{
    protected ExcelTransformer $transformer;

    protected function setUp(): void
    {
        parent::setUp();
        $this->transformer = new ExcelTransformer();
    }

    public function test_subject_code_normalization(): void
    {
        $testCases = [
            'T2811' => 'DSPD',
            'T2340' => 'DSPN',
            'T2320' => 'DSPF',
            'T2300' => 'DSPR',
            'T2310' => 'DSPL',
            'UT6522002' => 'SKEEH',
            'UT6521001' => 'SKMJH',
            'UT6521005' => 'SMJMH',
            'UT6523002' => 'SKEMH',
            'UT6523003' => 'SKEBH',
            'UT6481007' => 'SAIAH',
            'OTHER_CODE' => 'OTHER_CODE',
        ];

        foreach (['PUPW', 'Foundation'] as $sheetName) {
            foreach ($testCases as $input => $expected) {
                $row = ['ASA_KURTAWAR' => $input];
                $transformed = $this->transformer->transform($row, $sheetName);
                $this->assertEquals($expected, $transformed['ASA_KURTAWAR'], "Sheet: {$sheetName}, Input: {$input}");
            }
        }
    }

    public function test_subject_code_unchanged_for_utm_idp_sheet(): void
    {
        $testCases = [
            'T2811',
            'T2340',
            'UT6522002',
            'OTHER_CODE',
        ];

        foreach ($testCases as $input) {
            $row = ['ASA_KURTAWAR' => $input];
            $transformed = $this->transformer->transform($row, 'UTM-IDP');
            $this->assertSame($input, $transformed['ASA_KURTAWAR'], "Input: {$input}");
        }
    }

    public function test_kaum_to_kodtaraf(): void
    {
        $testCases = [
            'A' => 'B',
            'D' => 'B',
            'E' => 'B',
            'F' => 'B',
            'a' => 'B', // should trim and uppercase
            'd ' => 'B',
            'C' => '',
            'B' => '',
            'OTHER' => '',
        ];

        foreach ($testCases as $input => $expected) {
            $row = [
                'ASA_KAUM' => $input,
                'ASA_WARGA' => 'M'
            ];
            $transformed = $this->transformer->transform($row);
            $this->assertEquals($expected, $transformed['ASA_KODTARAF']);
        }
    }

    public function test_date_normalization(): void
    {
        // Excel serial date 45000 is 2023-03-15
        $row1 = ['ASA_TKHDAFTAR' => '45000', 'ASA_TKHLAHIR' => '25/12/1995'];
        $transformed1 = $this->transformer->transform($row1);
        $this->assertEquals('2023-03-15', $transformed1['ASA_TKHDAFTAR']);
        $this->assertEquals('1995-12-25', $transformed1['ASA_TKHLAHIR']);

        // Various string formats
        $testCases = [
            '2023-10-31' => '2023-10-31',
            '31-10-2023' => '2023-10-31',
            '31/10/2023' => '2023-10-31',
            'invalid-date' => '',
        ];

        foreach ($testCases as $input => $expected) {
            $row = ['ASA_TKHDAFTAR' => $input];
            $transformed = $this->transformer->transform($row);
            $this->assertEquals($expected, $transformed['ASA_TKHDAFTAR']);
        }
    }

    public function test_phone_number_preservation(): void
    {
        $testCases = [
            // Symbol stripping
            '+60 12-345 6789'  => '60123456789',
            '(012) 3456789'    => '0123456789',
            // Leading zero preserved
            '0123456789'       => '0123456789',
            // 12-digit stays intact (under 13-char limit)
            '628123456789'     => '628123456789',
            // Float from PhpSpreadsheet (scientific notation on read)
            60123456789.0      => '60123456789',
            628123456789.0     => '628123456789',
            // Empty
            ''                 => '',
        ];

        foreach ($testCases as $input => $expected) {
            $row = ['ASA_TELEFON' => $input];
            $transformed = $this->transformer->transform($row);
            $this->assertIsString($transformed['ASA_TELEFON'], 'ASA_TELEFON must be string type');
            $this->assertEquals($expected, $transformed['ASA_TELEFON'], 'Input: ' . var_export($input, true));
        }
    }

    public function test_name_cleaning(): void
    {
        $testCases = [
            "amir@amin, dato'" => "AMIR AMIN DATO",
            "dato', a'lia @ amin" => "DATO ALIA AMIN",
            "amir@amin" => "AMIR AMIN",
            "Mohd. Ali" => "MOHD ALI",
            "Double   Space" => "DOUBLE SPACE",
        ];

        foreach ($testCases as $input => $expected) {
            $row = ['ASA_NAMA' => $input];
            $transformed = $this->transformer->transform($row);
            $this->assertEquals($expected, $transformed['ASA_NAMA']);
        }
    }

    public function test_foreigner_detection(): void
    {
        // 1. Malaysian applicant - should NOT clear ASA_NEGARA
        $row1 = [
            'ASA_NOKP' => '990101-01-1234',
            'ASA_TELEFON' => '012-3456789',
            'ASA_NEGARA' => 'MALAYSIA',
        ];
        $transformed1 = $this->transformer->transform($row1);
        $this->assertEquals('MALAYSIA', $transformed1['ASA_NEGARA']);

        // 2. Foreigner applicant with ASA_NEGARA = MAL - should clear to ""
        $row2 = [
            'ASA_NOKP' => 'A1234567',
            'ASA_ALAMAT1' => 'JAKARTA INDONESIA',
            'ASA_NOHP' => '628123456789',
            'ASA_NEGARA' => 'MAL',
        ];
        $transformed2 = $this->transformer->transform($row2);
        $this->assertEquals('', $transformed2['ASA_NEGARA']);

        // 3. Foreigner applicant with ASA_NEGARA = SINGAPORE - should NOT clear
        $row3 = [
            'ASA_NOKP' => '12345678',
            'ASA_TELEFON' => '+6591234567',
            'ASA_NEGARA' => 'SINGAPORE',
        ];
        $transformed3 = $this->transformer->transform($row3);
        $this->assertEquals('SINGAPORE', $transformed3['ASA_NEGARA']);

        // 4. Foreigner applicant due to foreign address, even with valid IC/phone, with ASA_NEGARA = MAL - should clear to ""
        $row4 = [
            'ASA_NOKP' => '990101-01-1234',
            'ASA_TELEFON' => '0123456789',
            'ASA_ALAMAT1' => 'SINGAPORE',
            'ASA_NEGARA' => 'MAL',
        ];
        $transformed4 = $this->transformer->transform($row4);
        $this->assertEquals('', $transformed4['ASA_NEGARA']);
    }

    public function test_address_splitting(): void
    {
        config(['excel_rules.alamat1_max_length' => 40]);
        $row1 = [
            'ASA_ALAMAT1' => "No 13 18/5\nThe Starits View Home\nBandar Baru Permas Jaya",
            'ASA_ALAMAT2' => ''
        ];
        $transformed1 = $this->transformer->transform($row1);
        $this->assertEquals("No 13 18/5\nThe Starits View Home", $transformed1['ASA_ALAMAT1']);
        $this->assertEquals('Bandar Baru Permas Jaya', $transformed1['ASA_ALAMAT2']);
    }

    public function test_city_normalization(): void
    {
        $testCases = [
            'BANDAR PUTRA,KULAI' => 'BANDAR PUTRA',
            'PUCHONG, SELANGOR' => 'PUCHONG',
            'BANDAR BARU PERMAS JAYA' => 'BB PERMAS JAYA',
        ];

        foreach ($testCases as $input => $expected) {
            $row = ['ASA_BANDAR' => $input];
            $transformed = $this->transformer->transform($row);
            $this->assertEquals($expected, $transformed['ASA_BANDAR']);
        }
    }
}
