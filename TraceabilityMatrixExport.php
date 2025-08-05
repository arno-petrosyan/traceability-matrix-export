<?php

namespace App\Exports;

use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithStyles;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Events\AfterSheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class TraceabilityMatrixExport implements FromCollection, WithHeadings, WithStyles, WithEvents, ShouldAutoSize
{
    protected array $matrix;

    public function __construct(array $matrix)
    {
        $this->matrix = $matrix;
    }

    public function collection(): Collection
    {
        return collect($this->matrix)->map(function ($rowGroup) {
            $row = [];
            for ($i = 1; $i <= 5; $i++) {
                if (!empty($rowGroup[$i])) {
                    $item = $rowGroup[$i][0];
                    $row = [...$row, $item->prefix, $item->description];
                } else {
                    $row[] = '-';
                }
            }
            return $row;
        });
    }

    public function headings(): array
    {
        return [
            'User Need ID',
            'User Need Description',
            'Design Input ID',
            'Design Input Description',
            'Design Output ID',
            'Design Output Description',
            'Verification ID',
            'Verification Description',
            'Validation ID',
            'Validation Description'
        ];
    }

    public function styles(Worksheet $sheet): array
    {
        return [
            1 => ['font' => ['bold' => true]]
        ];
    }

    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function (AfterSheet $event) {
                $sheet = $event->sheet->getDelegate();
                $rowCount = count($this->matrix) + 1; // +1 for heading
                $columnCount = count($this->headings());

                $columnRange = range('A', chr(ord('A') + $columnCount - 1));
                $range = "A1:{$columnRange[$columnCount - 1]}{$rowCount}";

                // Apply thin borders to all cells
                $sheet->getStyle($range)->applyFromArray([
                    'borders' => [
                        'allBorders' => [
                            'borderStyle' => Border::BORDER_THIN,
                            'color' => ['argb' => 'FF000000'],
                        ],
                    ],
                ]);

                $sheet->getStyle("A1:{$columnRange[$columnCount - 1]}1")->applyFromArray([
                    'borders' => [
                        'allBorders' => [
                            'borderStyle' => Border::BORDER_MEDIUM,
                            'color' => ['argb' => 'FF000000'],
                        ],
                    ],
                ]);
            },
        ];
    }
}
