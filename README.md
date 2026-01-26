# Website Scraper

A collection of Python web scrapers for extracting prices from various websites.

## Scrapers

### 1. Gold Price Scraper
Extracts gold prices from [qatar-goldprice.com](https://qatar-goldprice.com/).

### 2. Hotel Price Scraper
Extracts hotel prices from [booking.com](https://www.booking.com) for hotels in Doha, Qatar.

### 3. Flight Price Scraper
Extracts flight ticket prices from multiple airlines and travel websites for weekly price tracking.

## Features

### Gold Price Scraper
- Scrapes gold prices for karats: 14, 18, 21, 22, 24
- Extracts prices in both Qatari Riyal (QAR) and US Dollar (USD)
- Saves results to JSON file
- Exports data to Excel file with formatted table (similar to website display)
- Displays formatted output in terminal
- Tracks prices over time (adds new date columns to Excel on each run)

### Hotel Price Scraper
- Scrapes prices for 33+ hotels in Doha, Qatar
- Uses Selenium for dynamic content scraping
- Saves results to JSON file
- Exports data to Excel file with weekly tracking
- Tracks prices weekly (adds new week columns to Excel)

### Flight Price Scraper
- Scrapes flight prices from multiple airlines and travel aggregators
- Tracks routes (e.g., Doha-London-Doha, Doha-Cairo-Doha)
- Filters for non-stop flights, Economy class, Flexible/Semi-flexible tickets
- Uses Selenium for dynamic content scraping
- Saves results to JSON file
- Exports data to Excel file with weekly tracking matching the required format
- Tracks prices weekly (adds new week columns to Excel)

## Setup

### Creating a Virtual Environment

It's recommended to use a virtual environment to isolate project dependencies:

1. Create a virtual environment:

```bash
python3 -m venv venv
```

2. Activate the virtual environment:

   **On Linux/macOS:**

   ```bash
   source venv/bin/activate
   ```

   **On Windows:**

   ```bash
   venv\Scripts\activate
   ```

3. Once activated, you'll see `(venv)` in your terminal prompt.

### Deactivating the Virtual Environment

To deactivate the virtual environment when you're done:

```bash
deactivate
```

## Installation

1. Make sure your virtual environment is activated (see Setup above)
2. Install required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Manual Run

Run the scraper manually:

```bash
python gold_scraper.py
```

The script will:

1. Fetch the webpage
2. Extract gold prices for the specified karats
3. Display formatted output in terminal
4. Save results to `gold_prices.json`
5. Export data to `gold_prices.xlsx` (Excel format)

### Daily Automatic Run

To run the scraper automatically every day, see **[SCHEDULER_SETUP.md](SCHEDULER_SETUP.md)** for detailed instructions.

**Quick start with Python scheduler:**

```bash
python scheduler.py
```

This will run the scraper immediately and then daily at 9:00 AM. You can customize the schedule in `scheduler.py`.

**Other options:**
- **Linux/macOS**: Use cron jobs (see SCHEDULER_SETUP.md)
- **Windows**: Use Task Scheduler (see SCHEDULER_SETUP.md)

### Excel Output Format

The Excel file (`gold_prices.xlsx`) is formatted as a right-to-left (RTL) table with:

- **Row 1**: Arabic day names (e.g., "الخميس", "الجمعة") in yellow header
- **Row 2**: Month and day number (e.g., "يناير 1") in yellow header, with "نوع العيار" in column A
- **Rows 3-7**: Gold prices for each karat (14K, 18K, 21K, 22K, 24K)
- **Column A**: Karat labels (14, 18, 21, 22, 24) with orange/peach background
- **Columns B+**: Price data for each date

Each time you run the scraper, it will:

- Add a new column with today's date if it doesn't already exist
- Update prices for the current date
- Preserve historical data from previous runs
- Display in RTL (right-to-left) layout for Arabic text

## Output

The scraper outputs:

- Gold prices per gram for each karat (14K, 18K, 21K, 22K, 24K)
- Prices in both QAR and USD
- Timestamp of when the data was scraped

## Example Output

```
============================================================
GOLD PRICES FROM QATAR-GOLDPRICE.COM
============================================================
Timestamp: 2026-01-16T10:30:00
Source: https://qatar-goldprice.com/

Price per Gram:
------------------------------------------------------------
  14K Gold:
    QAR: 314.48
    USD: 86.40

  18K Gold:
    QAR: 404.33
    USD: 111.08

  21K Gold:
    QAR: 471.72
    USD: 129.59

  22K Gold:
    QAR: 494.18
    USD: 135.76

  24K Gold:
    QAR: 539.11
    USD: 148.11
============================================================
```

## Hotel Price Scraper

### Usage

Run the hotel scraper:

```bash
python hotel_scraper.py
```

The script will:
1. Search for each hotel on booking.com
2. Extract current prices
3. Display results in terminal
4. Save results to `hotel_prices.json`
5. Export data to `hotel_prices.xlsx` (Excel format) with weekly tracking

### Weekly Automatic Run

To run the scraper automatically every week:

```bash
python hotel_scheduler.py
```

This will run the scraper weekly on Monday at 9:00 AM.

### Excel Output Format

The Excel file (`hotel_prices.xlsx`) is formatted as a right-to-left (RTL) table with:
- **Row 1**: Week label (e.g., "أسبوع 2025-01-13")
- **Row 2**: Date in yellow header, with "الفندق" (Hotel) in column A
- **Rows 3+**: Hotel names in column A with prices in date columns
- Each week adds a new column with that week's prices

## Notes

### Gold Price Scraper
- The scraper uses a User-Agent header to avoid being blocked
- Website structure may change, requiring updates to the parsing logic
- Prices are updated automatically on the website several times daily

### Hotel Price Scraper
- Booking.com has anti-scraping measures, so results may vary
- The scraper includes delays between requests to avoid being blocked
- Some hotels may not be found if names don't match exactly
- Prices are for a default date (7 days from today, 1 night stay, 2 adults)
- You may need to adjust search parameters in the code for better results

## Flight Price Scraper

### Usage

Run the flight scraper:

```bash
python flight_scraper.py
```

The script will:
1. Search for flights on multiple airlines and travel websites
2. Extract current prices for defined routes
3. Display results in terminal
4. Save results to `flight_prices.json`
5. Export data to `flight_prices.xlsx` (Excel format) with weekly tracking

### Weekly Automatic Run

To run the scraper automatically every week:

```bash
python flight_scheduler.py
```

This will run the scraper weekly on Monday at 9:00 AM.

### Excel Output Format

The Excel file (`flight_prices.xlsx`) is formatted as a right-to-left (RTL) table matching the required format with:
- **Column A**: Code (route identifier)
- **Column B**: Commodity (Arabic description of route and ticket type)
- **Column C**: Class (Economy)
- **Column D**: CPI-Flag (Y for individual sources, N-averages for averages)
- **Column E**: Source Code (e.g., AIRL001, AIRL018)
- **Column F**: Flight Agencies (airline/aggregator name in Arabic)
- **Columns 7+**: Weekly date columns (e.g., 3-Jan, 10-Jan, 17-Jan) with prices

The scraper tracks:
- Individual prices from each source
- Average prices grouped by airline
- Multiple routes (configurable in the code)

### Supported Airlines and Travel Websites

The scraper supports the following sources:

**Airlines:**
- Qatar Airways
- British Airways
- Malaysia Airlines
- Kuwait Airways
- Turkish Airlines
- Pakistan International Airlines (PIA)

**Travel Aggregators:**
- CheapAir
- eDreams
- KAYAK
- ITA Matrix

### Scraping Criteria

The scraper follows these criteria for data collection:
- **Same airline for consistency**: Prices are collected from the same airline when possible
- **Non-stop flights only**: Only direct flights are selected
- **Economy Class**: Only Economy Class fares are considered
- **Flexible/Semi-Flexible tickets**: Only flexible and semi-flexible ticket options are used
- **Weekly collection**: Data is collected on a weekly basis

### Notes

- Flight scraping is complex and each website has different structures
- The current implementation provides a framework that can be extended for each airline/website
- Some airlines may require specific implementations based on their website structure
- Anti-scraping measures may affect results - delays are included between requests
- Website structures may change over time, requiring updates to the parsing logic
- You may need to customize the scraping methods for each airline based on their specific website structure
