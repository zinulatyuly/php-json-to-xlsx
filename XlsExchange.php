<?php
require './vendor/autoload.php';

class XlsExchange {
    protected string $path_to_input_json_file = '';
    protected string $path_to_output_xlsx_file = '';
    protected string $ftp_host = '';
    protected string $ftp_login = '';
    protected string $ftp_password = '';
    protected string $ftp_dir = '';

    public function export(): void
    {
        try {
            $this->validate();
            $inputData = $this->parseInputFile();
            $this->storeXlsx($inputData);
            if ($this->ftp_host) $this->moveFileToFTP();
            echo 'Done!' . PHP_EOL;
        } catch (Exception $e) {
            exit("Error: " . $e->getMessage() . PHP_EOL);
        }
    }

    public function setInputFile(string $path): self
    {
        $this->path_to_input_json_file = $path;
        return $this;
    }

    public function setOutputFile(string $path): self
    {
        $this->path_to_output_xlsx_file = $path;
        return $this;
    }

    public function setFTP(string $host, string $login, string $password, string $dir = ''): self
    {
        if (!$host) throw new InvalidArgumentException('Host is required.');
        $this->ftp_host = $host;
        $this->ftp_login = $login;
        $this->ftp_password = $password;
        if ($dir) $this->ftp_dir = $dir;
        return $this;
    }

    protected function validate()
    {
        if (!$this->path_to_input_json_file) throw new Exception('Path to input file is required.');
        if (!file_exists($this->path_to_input_json_file)) throw new Exception("Input file doesn't exist.");
        if (!$this->path_to_output_xlsx_file) throw new Exception('Path to output file is required.');
    }

    protected function parseInputFile(): array
    {
        $input = file_get_contents($this->path_to_input_json_file);
        if (!$input) throw new Exception("Input file has no data.");

        $data = json_decode($input, true);
        // EAN-13
        foreach ($data['items'] as $item) {
            if (!preg_match("/^[0-9]{13}$/", $item['item']['barcode'])) {
                throw new Exception("Barcode is incorrect for item with id = {$item['item']['id']}.");
            }
        }

        return $data;
    }

    protected function storeXlsx(array $data): void {
        $settings = [
            [
                'name' => 'Id',
                'value' => function ($item): string { return $item['id']; },
            ],
            [
                'name' => 'ШК',
                'value' => function ($item): string { return $item['item']['barcode']; },
            ],
            [
                'name' => 'Название',
                'value' => function ($item): string { return $item['item']['name']; },
            ],
            [
                'name' => 'Кол-во',
                'value' => function ($item): string { return $item['quantity']; },
            ],
            [
                'name' => 'Сумма',
                'value' => function ($item): string { return $item['amount']; },
            ],
        ];

        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $currentRow = 1;
        foreach ($settings as $i => $s) {
            $columnIndex = $i + 1;
            $sheet->setCellValueByColumnAndRow($columnIndex, $currentRow, $s['name']);
            $sheet->getStyleByColumnAndRow($columnIndex, $currentRow)->getFont()->setBold(true);
            $sheet->getStyleByColumnAndRow($columnIndex, $currentRow)->getAlignment()->setHorizontal('center');
            $sheet->getColumnDimensionByColumn($columnIndex)->setAutoSize(true);
        }
        $currentRow++;

        foreach ($data['items'] as $index => $item) {
            foreach ($settings as $i => $s) {
                $columnIndex = $i + 1;
                $sheet->setCellValueByColumnAndRow($columnIndex, $currentRow + $index, $s['value']($item));
            }
        }

        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save($this->path_to_output_xlsx_file);
    }

    protected function moveFileToFTP()
    {
        $path = $this->path_to_output_xlsx_file;
        $outputPath = $this->ftp_dir . DIRECTORY_SEPARATOR . $path;
        $connection = ftp_connect($this->ftp_host) or die("Can't connect to FTP server." . PHP_EOL);

        try {
            $checkCredentials = @ftp_login($connection, $this->ftp_login, $this->ftp_password);
            if (!$checkCredentials) throw new Exception('Wrong FTP credentials.');

            if (!ftp_put($connection, $outputPath, $path, FTP_ASCII)) throw new Exception("Can't send file to FTP server.");
        } catch (Exception $e) {
            throw $e;
        } finally {
            ftp_close($connection);
            unlink($this->path_to_output_xlsx_file);
        }
    }
}