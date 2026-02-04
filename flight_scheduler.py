#!/usr/bin/env python3
"""
Monthly scheduler for flight price scraper
Runs the scraper on the 4th, 10th, 17th, and 24th of each month at 9:00 AM
"""

import schedule
import time
from datetime import datetime
from flight_scraper import FlightPriceScraper
import json


def run_flight_scraper():
    """Run the flight price scraper"""
    print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Running flight price scraper...")
    try:
        scraper = FlightPriceScraper(headless=True)
        results = scraper.scrape_all()
        
        # Save to JSON
        json_file = 'flight_prices.json'
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"Prices saved to {json_file}")
        
        # Export to Excel
        excel_file = 'flight_prices.xlsx'
        scraper.export_to_excel(results, excel_file)
        
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Scraper completed successfully\n")
    except Exception as e:
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Error running scraper: {e}\n")


def should_run_today():
    """Check if today is one of the scheduled days (4th, 10th, 17th, 24th)"""
    today = datetime.now()
    scheduled_days = [4, 10, 17, 24]
    return today.day in scheduled_days


def check_and_run():
    """Check if today is a scheduled day and run the scraper"""
    if should_run_today():
        # Check if we've already run today (to avoid multiple runs)
        today_str = datetime.now().strftime('%Y-%m-%d')
        try:
            with open('last_run_date.txt', 'r') as f:
                last_run = f.read().strip()
            if last_run == today_str:
                print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Already ran today, skipping...")
                return
        except FileNotFoundError:
            pass
        
        # Run the scraper
        run_flight_scraper()
        
        # Save the run date
        with open('last_run_date.txt', 'w') as f:
            f.write(today_str)


def main():
    """Main scheduler function"""
    # Schedule to check every day at 9:00 AM
    schedule.every().day.at("09:00").do(check_and_run)
    
    print("Flight Price Scraper Scheduler Started")
    print("=" * 60)
    print("Scheduled to run on the 4th, 10th, 17th, and 24th of each month at 09:00")
    print("Press Ctrl+C to stop the scheduler")
    print("=" * 60)
    
    # Check if we should run immediately
    now = datetime.now()
    if should_run_today() and now.hour >= 9:
        # If it's a scheduled day and past 9 AM, run immediately
        check_and_run()
    elif should_run_today() and now.hour < 9:
        print(f"Today ({now.day}) is a scheduled day. Will run at 09:00")
    else:
        next_scheduled = min([d for d in [4, 10, 17, 24] if d > now.day], default=None)
        if next_scheduled:
            print(f"Next scheduled run: {next_scheduled} of this month at 09:00")
        else:
            # Next month
            print(f"Next scheduled run: 4th of next month at 09:00")
    
    # Keep the script running
    while True:
        schedule.run_pending()
        time.sleep(60)  # Check every minute


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nScheduler stopped by user")
