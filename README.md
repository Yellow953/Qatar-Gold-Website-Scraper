# Gold Price Scraper

A Python web scraper to extract gold prices from [qatar-goldprice.com](https://qatar-goldprice.com/).

## Features

- Scrapes gold prices for karats: 14, 18, 21, 22, 24
- Extracts prices in both Qatari Riyal (QAR) and US Dollar (USD)
- Saves results to JSON file
- Exports data to Excel file with formatted table (similar to website display)
- Displays formatted output in terminal
- Tracks prices over time (adds new date columns to Excel on each run)

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

Run the scraper:

```bash
python gold_scraper.py
```

The script will:

1. Fetch the webpage
2. Extract gold prices for the specified karats
3. Display formatted output in terminal
4. Save results to `gold_prices.json`
5. Export data to `gold_prices.xlsx` (Excel format)

### Excel Output Format

The Excel file (`gold_prices.xlsx`) is formatted as a table with:

- **Row 1**: Title "الذهب 2025" (Gold 2025) in green header
- **Row 2**: Date headers in yellow (Arabic day names and dates)
- **Rows 3-7**: Gold prices for each karat (14K, 18K, 21K, 22K, 24K)
- **Column A**: Karat labels (14K, 18K, etc.)
- **Columns B+**: Price data for each date

Each time you run the scraper, it will:

- Add a new column with today's date if it doesn't already exist
- Update prices for the current date
- Preserve historical data from previous runs

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

## Notes

- The scraper uses a User-Agent header to avoid being blocked
- Website structure may change, requiring updates to the parsing logic
- Prices are updated automatically on the website several times daily
