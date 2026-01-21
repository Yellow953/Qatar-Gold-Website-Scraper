#!/usr/bin/env python3
"""
Weekly scheduler for hotel price scraper
Runs the scraper weekly at a specified time
"""

import schedule
import time
from datetime import datetime
from hotel_scraper import HotelPriceScraper
import json


def run_hotel_scraper():
    """Run the hotel price scraper"""
    print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Running hotel price scraper...")
    try:
        scraper = HotelPriceScraper()
        results = scraper.scrape_all_hotels()
        
        # Save to JSON
        json_file = 'hotel_prices.json'
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"Prices saved to {json_file}")
        
        # Export to Excel
        excel_file = 'hotel_prices.xlsx'
        scraper.export_to_excel(results, excel_file)
        
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Scraper completed successfully\n")
    except Exception as e:
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Error running scraper: {e}\n")


def main():
    """Main scheduler function"""
    # Schedule the scraper to run weekly on Monday at 9:00 AM
    # You can change this time/day to any you prefer
    schedule.every().monday.at("09:00").do(run_hotel_scraper)
    
    # Optional: Run immediately on startup
    print("Hotel Price Scraper Scheduler Started")
    print("=" * 60)
    print("Scheduled to run weekly on Monday at 09:00")
    print("Press Ctrl+C to stop the scheduler")
    print("=" * 60)
    run_hotel_scraper()  # Run once immediately
    
    # Keep the script running
    while True:
        schedule.run_pending()
        time.sleep(3600)  # Check every hour (weekly schedule doesn't need minute checks)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nScheduler stopped by user")
