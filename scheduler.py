#!/usr/bin/env python3
"""
Daily scheduler for gold price scraper
Runs the scraper at a specified time each day
"""

import schedule
import time
from datetime import datetime
from gold_scraper import GoldPriceScraper


def run_scraper():
    """Run the gold price scraper"""
    print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Running gold price scraper...")
    try:
        scraper = GoldPriceScraper()
        prices = scraper.scrape()
        scraper.print_prices(prices)
        
        # Save to JSON
        import json
        json_file = 'gold_prices.json'
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(prices, f, indent=2, ensure_ascii=False)
        print(f"Prices saved to {json_file}")
        
        # Export to Excel
        excel_file = 'gold_prices.xlsx'
        scraper.export_to_excel(prices, excel_file)
        
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Scraper completed successfully\n")
    except Exception as e:
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Error running scraper: {e}\n")


def main():
    """Main scheduler function"""
    # Schedule the scraper to run daily at 9:00 AM
    # You can change this time to any time you prefer
    schedule.every().day.at("09:00").do(run_scraper)
    
    # Optional: Run immediately on startup
    print("Gold Price Scraper Scheduler Started")
    print("=" * 60)
    print("Scheduled to run daily at 09:00")
    print("Press Ctrl+C to stop the scheduler")
    print("=" * 60)
    run_scraper()  # Run once immediately
    
    # Keep the script running
    while True:
        schedule.run_pending()
        time.sleep(60)  # Check every minute


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nScheduler stopped by user")
