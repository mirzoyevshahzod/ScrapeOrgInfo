<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use GuzzleHttp\Client;
use Symfony\Component\DomCrawler\Crawler;

class ScrapeCompanyNames extends Command
{
    protected $signature = 'scrape:companies';
    protected $description = 'Scrape company names by INN numbers and export to Excel';

    public function handle()
    {
        $this->info('Starting the scraping process...');

        $filePath = storage_path('app/inn_numbers.xlsx');
        $spreadsheet = IOFactory::load($filePath);
        $worksheet = $spreadsheet->getActiveSheet();
        $innNumbers = [];

        foreach ($worksheet->getColumnIterator('A') as $column) {
            foreach ($column->getCellIterator() as $cell) {
                $innNumbers[] = $cell->getValue();
            }
        }

        $client = new Client(['verify' => false]);
        $scrapedData = [];

        foreach ($innNumbers as $inn) {
            $this->info("Processing INN: $inn");

            try {
                $response = $client->get('https://orginfo.uz/', [
                    'query' => ['q' => $inn],
                    'headers' => [
                        'User-Agent' => 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                    ],
                ]);

                $html = $response->getBody()->getContents();

                // Save HTML for debugging purposes
                file_put_contents(storage_path("app/html_debug_$inn.html"), $html);

                $crawler = new Crawler($html);

                try {
                    $companyName = $crawler->filter('h1.h1-seo')->text();
                } catch (\Exception $e) {
                    $this->error("Company name not found for INN: $inn");
                    $companyName = 'Not Found';
                }

                $scrapedData[] = [
                    'INN' => $inn,
                    'Company Name' => $companyName,
                ];
            } catch (\Exception $e) {
                $this->error("Failed to fetch data for INN: $inn");
                $scrapedData[] = [
                    'INN' => $inn,
                    'Company Name' => 'Not Found',
                ];
            }
        }

        $this->exportToExcel($scrapedData);
        $this->info('Scraping completed successfully.');
    }

    private function exportToExcel(array $data)
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $sheet->setCellValue('A1', 'INN');
        $sheet->setCellValue('B1', 'Company Name');

        foreach ($data as $index => $row) {
            $sheet->setCellValue('A' . ($index + 2), $row['INN']);
            $sheet->setCellValue('B' . ($index + 2), $row['Company Name']);
        }

        $filePath = storage_path('app/scraped_companies.xlsx');
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($filePath);

        $this->info("Data exported to: $filePath");
    }
}
