<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Facebook\WebDriver\Remote\RemoteWebDriver;
use Facebook\WebDriver\Remote\DesiredCapabilities;
use Facebook\WebDriver\Chrome\ChromeOptions;
use Facebook\WebDriver\WebDriverBy;

class ScrapeCompanyNames extends Command
{
    protected $signature = 'scrape:companies';
    protected $description = 'Scrape company names by INN numbers and export to Excel';

    public function handle()
    {
        $this->info('Starting the scraping process...');

        // Load INN numbers from an Excel file
        $filePath = storage_path('app/inn_numbers.xlsx'); // Replace with your file path
        $spreadsheet = IOFactory::load($filePath);
        $worksheet = $spreadsheet->getActiveSheet();
        $innNumbers = [];

        foreach ($worksheet->getColumnIterator('A') as $column) {
            foreach ($column->getCellIterator() as $cell) {
                $innNumbers[] = $cell->getValue();
            }
        }

        // Configure Selenium
        $serverUrl = 'http://localhost:4444/wd/hub'; // Adjust this URL if necessary
        $options = new ChromeOptions();
        $options->addArguments(['--headless', '--disable-gpu', '--window-size=1920,1080']);
        $capabilities = DesiredCapabilities::chrome()->setCapability(ChromeOptions::CAPABILITY, $options);
        $driver = RemoteWebDriver::create($serverUrl, $capabilities);

        $scrapedData = [];

        foreach ($innNumbers as $inn) {
            $this->info("Processing INN: $inn");

            try {
                // Navigate to the target website
                $driver->get('https://orginfo.uz/');

                // Input the INN into the search field
                $searchInput = $driver->findElement(WebDriverBy::name('q'));
                $searchInput->sendKeys($inn);

                // Submit the search
                $searchInput->submit();

                // Wait for the results to load
                sleep(2);

                // Extract the company name
                $companyNameElement = $driver->findElement(WebDriverBy::cssSelector('h1.h1-seo'));
                $companyName = $companyNameElement->getText();

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

        // Quit the WebDriver
        $driver->quit();

        // Export the scraped data to a new Excel file
        $this->exportToExcel($scrapedData);

        $this->info('Scraping completed successfully.');
    }

    private function exportToExcel(array $data)
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // Write header row
        $sheet->setCellValue('A1', 'INN');
        $sheet->setCellValue('B1', 'Company Name');

        // Write data rows
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
