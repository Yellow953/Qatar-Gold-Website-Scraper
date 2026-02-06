#!/usr/bin/env python3
"""
Flight Ticket Price Scraper
Extracts flight prices from multiple airlines and travel websites
for weekly price tracking

IMPORTANT NOTES FOR IMPLEMENTATION:
===================================

This scraper provides a framework for scraping flight prices. Each airline/website
has a different structure and may require specific implementation:

1. **Implementing Airline Scrapers:**
   - Each airline website (Qatar Airways, British Airways, etc.) has unique HTML structure
   - You'll need to inspect each website to find the correct selectors for:
     * Origin/destination input fields
     * Date pickers
     * Class selection (Economy)
     * Ticket type selection (Flexible/Semi-flexible)
     * Non-stop flight filter
     * Price extraction
   
2. **Implementing Aggregator Scrapers:**
   - Aggregators like KAYAK, eDreams, etc. may be easier to scrape
   - They often have more consistent structures
   - You may need to filter results by airline name
   
3. **Testing:**
   - Test each scraper individually before running the full suite
   - Websites may change their structure, requiring updates
   - Some sites have anti-scraping measures - use delays and proper headers

4. **Data Requirements:**
   - Only non-stop (direct) flights
   - Economy class only
   - Flexible or Semi-flexible tickets only
   - Prices should be in QAR (Qatari Riyal) when possible

5. **Routes (add/delete in Excel):**
   - Routes are read from the top of the single sheet "Flight Prices" (route block: row 1 = headers, rows 2+ = data).
   - Columns: Code, Commodity, Origin, Origin_Code, Destination, Destination_Code, Duration_Months.
   - Add or delete rows in that block to change which routes are scraped.
"""

try:
    import undetected_chromedriver as uc
    USE_UNDETECTED = True
except ImportError:
    USE_UNDETECTED = False

import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import json
import time
import base64
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import requests
from bs4 import BeautifulSoup


# Excel file name and sheet names (single sheet: routes at top, flight prices below)
FLIGHT_PRICES_EXCEL = 'flight_prices.xlsx'
FLIGHT_PRICES_SHEET_NAME = 'Flight Prices'
ROUTE_HEADERS = ['Code', 'Commodity', 'Origin', 'Origin_Code', 'Destination', 'Destination_Code', 'Duration_Months']
FLIGHT_HEADER_MARKER = 'وكالات'  # Part of "وكالات الخطوط (Flight Agencies)" to find flight header row


class FlightPriceScraper:
    """Scraper for flight prices from multiple airlines and travel websites.
    Routes are read from the Excel file (sheet 'Routes'). Add or delete rows there to change what is scraped."""
    
    def __init__(self, headless=False, excel_path: str = None):
        self.headless = headless
        self.driver = None
        self.excel_path = excel_path or FLIGHT_PRICES_EXCEL
        self.routes = self._get_routes()
        self.sources = self._get_sources()
        self.results = []
    
    def _load_routes_from_excel(self) -> Optional[List[Dict]]:
        """Load routes from the top of the single Excel sheet (route block: row 1 = headers, rows 2+ = data until blank or flight header)."""
        if not os.path.exists(self.excel_path):
            return None
        try:
            wb = load_workbook(self.excel_path, read_only=True, data_only=True)
            if FLIGHT_PRICES_SHEET_NAME not in wb.sheetnames:
                ws = wb.active  # fallback for old files with different sheet name
            else:
                ws = wb[FLIGHT_PRICES_SHEET_NAME]
            routes = []
            # Route block: row 1 = headers (Code, Commodity, Origin, ...), rows 2+ = data. Stop at empty A or flight header row.
            first_cell = ws.cell(row=1, column=1).value
            col3 = str(ws.cell(row=1, column=3).value or '')
            if not first_cell or str(first_cell).strip() != 'Code':
                wb.close()
                return None
            if 'Class' not in col3 and 'الدرجة' not in col3:
                for row_idx in range(2, ws.max_row + 1):
                    code = ws.cell(row=row_idx, column=1).value
                    if not code or not str(code).strip():
                        break
                    col6 = str(ws.cell(row=row_idx, column=6).value or '')
                    if FLIGHT_HEADER_MARKER in col6:
                        break
                    commodity_ar = ws.cell(row=row_idx, column=2).value or ''
                    origin = ws.cell(row=row_idx, column=3).value or 'Doha'
                    origin_code = ws.cell(row=row_idx, column=4).value or 'DOH'
                    destination = ws.cell(row=row_idx, column=5).value
                    destination_code = ws.cell(row=row_idx, column=6).value
                    duration = ws.cell(row=row_idx, column=7).value
                    if not destination or not destination_code:
                        continue
                    try:
                        duration_months = int(duration) if duration is not None else 6
                    except (TypeError, ValueError):
                        duration_months = 6
                    routes.append({
                        'code': str(code).strip(),
                        'origin': str(origin).strip(),
                        'origin_code': str(origin_code).strip().upper(),
                        'destination': str(destination).strip(),
                        'destination_code': str(destination_code).strip().upper(),
                        'commodity_ar': str(commodity_ar).strip(),
                        'commodity_en': f"Cost of a {origin} - {destination} - {origin} ticket for {duration_months} months",
                        'ticket_type': 'Semi flexible',
                        'duration_months': duration_months
                    })
            if routes:
                wb.close()
                return routes
            if 'Routes' in wb.sheetnames:
                ws_legacy = wb['Routes']
                routes = []
                for row_idx in range(2, ws_legacy.max_row + 1):
                    code = ws_legacy.cell(row=row_idx, column=1).value
                    if not code or not str(code).strip():
                        continue
                    commodity_ar = ws_legacy.cell(row=row_idx, column=2).value or ''
                    origin = ws_legacy.cell(row=row_idx, column=3).value or 'Doha'
                    origin_code = ws_legacy.cell(row=row_idx, column=4).value or 'DOH'
                    destination = ws_legacy.cell(row=row_idx, column=5).value
                    destination_code = ws_legacy.cell(row=row_idx, column=6).value
                    duration = ws_legacy.cell(row=row_idx, column=7).value
                    if not destination or not destination_code:
                        continue
                    try:
                        duration_months = int(duration) if duration is not None else 6
                    except (TypeError, ValueError):
                        duration_months = 6
                    routes.append({
                        'code': str(code).strip(),
                        'origin': str(origin).strip(),
                        'origin_code': str(origin_code).strip().upper(),
                        'destination': str(destination).strip(),
                        'destination_code': str(destination_code).strip().upper(),
                        'commodity_ar': str(commodity_ar).strip(),
                        'commodity_en': f"Cost of a {origin} - {destination} - {origin} ticket for {duration_months} months",
                        'ticket_type': 'Semi flexible',
                        'duration_months': duration_months
                    })
            wb.close()
            return routes if routes else None
        except Exception as e:
            print(f"  Note: Could not load routes from Excel ({e}), using defaults.")
            return None
    
    def _get_routes(self) -> List[Dict]:
        """Get list of flight routes: from Excel if available, otherwise defaults."""
        loaded = self._load_routes_from_excel()
        if loaded:
            print(f"  Loaded {len(loaded)} routes from Excel (top of sheet '{FLIGHT_PRICES_SHEET_NAME}')")
            return loaded
        print(f"  Using default routes (edit route block at top of '{self.excel_path}' to add/remove routes)")
        return self._get_default_routes()
    
    def _get_default_routes(self) -> List[Dict]:
        """Default route list (used when Excel has no Routes sheet or no data)."""
        routes = [
            {
                'code': '007331101',
                'origin': 'Doha',
                'origin_code': 'DOH',
                'destination': 'London',
                'destination_code': 'LHR',
                'commodity_ar': 'كلفة تذكرة دوحة _ لندن - دوحة لمدة 6 (Semi flexble التذكرة السياحية) أشهر',
                'commodity_en': 'Cost of a Doha - London - Doha ticket for 6 (Semi flexible tourist ticket) months',
                'ticket_type': 'Semi flexible',
                'duration_months': 6
            },
            {
                'code': '007331102',
                'origin': 'Doha',
                'origin_code': 'DOH',
                'destination': 'Cairo',
                'destination_code': 'CAI',
                'commodity_ar': 'كلفة تذكرة دوحة _ القاهرة - دوحة لمدة 6 (semi flexble التذكرة سياحية ( اشهر',
                'commodity_en': 'Cost of a Doha - Cairo - Doha ticket for 6 (semi flexible tourist ticket) months',
                'ticket_type': 'Semi flexible',
                'duration_months': 6
            },
            {
                'code': '007331103',
                'origin': 'Doha',
                'origin_code': 'DOH',
                'destination': 'Karachi',
                'destination_code': 'KHI',
                'commodity_ar': 'كلفة تذكرة دوحة_ كراتشي _ دوحة لمدة 3 اشهر ( التذكرة سياحية semi flexble)',
                'commodity_en': 'Cost of a Doha - Karachi - Doha ticket for 3 (semi flexible tourist ticket) months',
                'ticket_type': 'Semi flexible',
                'duration_months': 3
            },
            {
                'code': '007331104',
                'origin': 'Doha',
                'origin_code': 'DOH',
                'destination': 'Dubai',
                'destination_code': 'DXB',
                'commodity_ar': 'كلفة تذكرة دوحة_ دبي _ دوحة لمدة 6 اشهر ( التذكرة سياحية semi flexble)',
                'commodity_en': 'Cost of a Doha - Dubai - Doha ticket for 6 (semi flexible tourist ticket) months',
                'ticket_type': 'Semi flexible',
                'duration_months': 6
            },
            {
                'code': '007331105',
                'origin': 'Doha',
                'origin_code': 'DOH',
                'destination': 'Jeddah',
                'destination_code': 'JED',
                'commodity_ar': 'كلفة تذكرة دوحة_جدة _ دوحة لمدة 6 اشهر( التذكرة سياحية semi flexble)',
                'commodity_en': 'Cost of a Doha - Jeddah - Doha ticket for 6 (semi flexible tourist ticket) months',
                'ticket_type': 'Semi flexible',
                'duration_months': 6
            },
            {
                'code': '007331106',
                'origin': 'Doha',
                'origin_code': 'DOH',
                'destination': 'Mumbai',
                'destination_code': 'BOM',
                'commodity_ar': 'كلفة تذكرة دوحة_ بومباي _ دوحة لمدة 3 اشهر ( التذكرة سياحية semi flexble)',
                'commodity_en': 'Cost of a Doha - Mumbai - Doha ticket for 3 (semi flexible tourist ticket) months',
                'ticket_type': 'Semi flexible',
                'duration_months': 3
            },
            {
                'code': '007331107',
                'origin': 'Doha',
                'origin_code': 'DOH',
                'destination': 'Kuala Lumpur',
                'destination_code': 'KUL',
                'commodity_ar': 'كلفة تذكرة دوحة_كولا لمبور _ دوحة لمدة 6 اشهر( التذكرة سياحية semi flexble)',
                'commodity_en': 'Cost of a Doha - Kuala Lumpur - Doha ticket for 6 (semi flexible tourist ticket) months',
                'ticket_type': 'Semi flexible',
                'duration_months': 6
            },
            {
                'code': '007331108',
                'origin': 'Doha',
                'origin_code': 'DOH',
                'destination': 'Istanbul',
                'destination_code': 'IST',
                'commodity_ar': 'كلفة تذكرة دوحة_ اسطنبول لمدة 6 اشهر ( التذكرة سياحية semi flexble)',
                'commodity_en': 'Cost of a Doha - Istanbul - Doha ticket for 6 (semi flexible tourist ticket) months',
                'ticket_type': 'Semi flexible',
                'duration_months': 6
            },
            {
                'code': '007331109',
                'origin': 'Doha',
                'origin_code': 'DOH',
                'destination': 'Bangkok',
                'destination_code': 'BKK',
                'commodity_ar': 'كلفة تذكرة دوحة_ بانكوك _ دوحة لمدة 6 اشهر ( التذكرة سياحية semi flexble)',
                'commodity_en': 'Cost of a Doha - Bangkok - Doha ticket for 6 (semi flexible tourist ticket) months',
                'ticket_type': 'Semi flexible',
                'duration_months': 6
            },
            {
                'code': '007331110',
                'origin': 'Doha',
                'origin_code': 'DOH',
                'destination': 'Tbilisi',
                'destination_code': 'TBS',
                'commodity_ar': 'كلفة تذكرة دوحة_تبليسي_ دوحة لمدة 6 اشهر ( التذكرة سياحية semi flexble)',
                'commodity_en': 'Cost of a Doha - Tbilisi - Doha ticket for 6 (semi flexible tourist ticket) months',
                'ticket_type': 'Semi flexible',
                'duration_months': 6
            },
            {
                'code': '007331111',
                'origin': 'Doha',
                'origin_code': 'DOH',
                'destination': 'New York',
                'destination_code': 'JFK',
                'commodity_ar': 'كلفة تذكرة دوحة_نيويورك دوحة لمدة 6 اشهر ( التذكرة سياحية semi flexble)',
                'commodity_en': 'Cost of a Doha - New York - Doha ticket for 6 (semi flexible tourist ticket) months',
                'ticket_type': 'Semi flexible',
                'duration_months': 6
            }
        ]
        return routes
    
    def _get_sources(self) -> List[Dict]:
        """Get list of airlines and travel websites to scrape"""
        sources = [
            {
                'name': 'Qatar Airways',
                'name_ar': 'الخطوط القطرية',
                'url': 'https://www.qatarairways.com/en-qa/homepage.html',
                'source_code': 'AIRL001',
                'type': 'airline'
            },
            {
                'name': 'British Airways',
                'name_ar': 'الخطوط البريطانية',
                'url': 'https://www.britishairways.com/travel/home/public/en_us/',
                'source_code': 'AIRL018',
                'type': 'airline'
            },
            {
                'name': 'Malaysia Airlines',
                'name_ar': 'الخطوط الماليزية',
                'url': 'https://www.malaysiaairlines.com/qa/en/home.html',
                'source_code': 'AIRL024',
                'type': 'airline'
            },
            {
                'name': 'Kuwait Airways',
                'name_ar': 'الخطوط الكويتية',
                'url': 'https://www.kuwaitairways.com/en',
                'source_code': 'AIRL025',
                'type': 'airline'
            },
            {
                'name': 'Turkish Airlines',
                'name_ar': 'الخطوط التركية',
                'url': 'https://www.turkishairlines.com/en-qa',
                'source_code': 'AIRL026',
                'type': 'airline'
            },
            {
                'name': 'Pakistan International Airlines',
                'name_ar': 'الخطوط الباكستانية',
                'url': 'https://www.piac.com.pk',
                'source_code': 'AIRL020',
                'type': 'airline'
            },
            {
                'name': 'CheapAir',
                'name_ar': 'cheapair',
                'url': 'https://www.cheapoair.com',  # Note: actual domain is cheapoair.com
                'source_code': 'AIRL028',
                'type': 'aggregator'
            },
            {
                'name': 'eDreams',
                'name_ar': 'edreams',
                'url': 'https://www.edreams.qa/home/',
                'source_code': 'AIRL030',
                'type': 'aggregator'
            },
            {
                'name': 'KAYAK',
                'name_ar': 'Kayak',
                'url': 'https://www.kayak.ae/?ispredir=true',
                'source_code': 'AIRL028',
                'type': 'aggregator'
            },
            {
                'name': 'ITA Matrix',
                'name_ar': 'matrix',
                'url': 'https://matrix.itasoftware.com/search',
                'source_code': 'AIRL028',
                'type': 'aggregator'
            }
        ]
        return sources
    
    def _setup_driver(self, headless=False):
        """Setup Chrome WebDriver"""
        if USE_UNDETECTED:
            try:
                options = uc.ChromeOptions()
                if headless:
                    options.add_argument('--headless=new')
                options.add_argument('--no-sandbox')
                options.add_argument('--disable-dev-shm-usage')
                options.add_argument('--start-maximized')
                self.driver = uc.Chrome(options=options, version_main=None)
                if not headless:
                    self.driver.set_window_size(1920, 1080)
                print("    Using undetected-chromedriver")
                return
            except Exception as e:
                print(f"    Warning: Could not use undetected-chromedriver: {e}")
        
        # Fallback to standard Selenium
        chrome_options = Options()
        if headless:
            chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument('user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.driver.set_window_size(1920, 1080)
    
    def _close_driver(self):
        """Close the WebDriver"""
        if self.driver:
            self.driver.quit()
    
    def _calculate_dates(self, months_ahead: int = 6) -> Tuple[str, str]:
        """Flight search dates: departure = 1 week after execution, return = departure + duration (e.g. 6 months)."""
        today = datetime.now()
        departure_date = today + timedelta(days=7)   # 1 week after execution
        return_date = departure_date + timedelta(days=months_ahead * 30)  # e.g. 6 months later
        
        return departure_date.strftime('%Y-%m-%d'), return_date.strftime('%Y-%m-%d')
    
    # Scheduled run days each month (same as scheduler: 4th, 10th, 17th, 24th)
    SCHEDULED_DAYS = (4, 10, 17, 24)
    
    def _get_scheduled_dates_through_2026(self) -> List[Tuple[str, datetime]]:
        """Return list of (header_text, date) for 4th, 10th, 17th, 24th of each month from today through end of 2026."""
        from datetime import date
        today = date.today()
        result = []
        for year in range(today.year, 2027):
            for month in range(1, 13):
                if year == 2026 and month > 12:
                    break
                for day in self.SCHEDULED_DAYS:
                    try:
                        d = date(year, month, day)
                        if d >= today:
                            header = d.strftime('%d-%b')  # e.g. 04-Jan, 10-Jan
                            result.append((header, datetime.combine(d, datetime.min.time())))
                    except ValueError:
                        pass  # e.g. Feb 30
        return result
    
    def _get_current_run_date_header(self) -> str:
        """Return the date header for this run: the scheduled date (4/10/17/24) that is >= today."""
        scheduled = self._get_scheduled_dates_through_2026()
        if not scheduled:
            return datetime.now().strftime('%d-%b')
        return scheduled[0][0]
    
    def _extract_price_from_text(self, text: str) -> Optional[float]:
        """Extract numeric price from text"""
        try:
            # Remove currency symbols and extract numbers
            cleaned = re.sub(r'[^\d.,]', '', text)
            cleaned = cleaned.replace(',', '')
            price = float(cleaned)
            return price if price > 0 else None
        except:
            return None
    
    def _extract_price_from_page(self, selectors: List[str], min_price: int = 500, max_price: int = 50000) -> Optional[float]:
        """Extract price from page using multiple selectors.
        min_price=500 to avoid capturing wrong values (baggage, taxes, etc.) from Kuwait/Malaysia/PIA pages.
        International economy fares from Doha are typically 500+ QAR."""
        price = None
        
        # Try each selector
        for selector in selectors:
            try:
                price_elems = self.driver.find_elements(By.CSS_SELECTOR, selector)
                if price_elems:
                    # Try the first few price elements
                    for price_elem in price_elems[:10]:
                        try:
                            price_text = price_elem.text.strip()
                            if price_text:
                                extracted = self._extract_price_from_text(price_text)
                                if extracted and min_price <= extracted <= max_price:
                                    price = extracted
                                    print(f"      ✓ Found price: {price} using selector: {selector[:50]}")
                                    break
                        except:
                            continue
                    if price:
                        break
            except:
                continue
        
        # If no price found, try searching page source
        if not price:
            try:
                page_text = self.driver.page_source
                # Prefer patterns that match typical fare amounts (4-6 digits); avoid small numbers
                price_patterns = [
                    r'QAR\s*(\d{1,3}(?:,\d{3})*)',  # QAR followed by number
                    r'(\d{4,6})',   # 4-6 digit numbers (typical fare range, avoids 100/202/103)
                    r'(\d{1,3}(?:,\d{3})+(?:\.\d{2})?)',  # Numbers with commas (e.g. 1,234.56)
                    r'(\d{3,6})',   # 3-6 digit numbers (fallback)
                ]
                for pattern in price_patterns:
                    matches = re.findall(pattern, page_text)
                    if matches:
                        # Sort candidates by value and pick first in valid range (prefer lower fare)
                        candidates = []
                        for match in matches[:30]:
                            try:
                                price_val = float(str(match).replace(',', ''))
                                if min_price <= price_val <= max_price:
                                    candidates.append(price_val)
                            except:
                                continue
                        if candidates:
                            # Use the smallest valid price (likely the main fare, not total with extras)
                            price = min(candidates)
                            print(f"      ✓ Extracted price from page source: {price}")
                            break
            except:
                pass
        
        return price
    
    def scrape_qatar_airways(self, route: Dict) -> Optional[Dict]:
        """Scrape prices from Qatar Airways"""
        try:
            print(f"    Scraping Qatar Airways for {route['origin']}-{route['destination']}")
            dep_date, ret_date = self._calculate_dates(route['duration_months'])
            
            # Direct URL with search parameters (direct/non-stop flights only)
            url = (f"https://www.qatarairways.com/app/booking/flight-selection?"
                   f"widget=QR&searchType=F&addTaxToFare=Y&minPurTime=0&selLang=en&"
                   f"tripType=R&fromStation={route['origin_code']}&toStation={route['destination_code']}&"
                   f"departing={dep_date}&returning={ret_date}&bookingClass=E&"
                   f"adults=1&children=0&infants=0&ofw=0&teenager=0&flexibleDate=off&allowRedemption=N&stops=0")
            
            print(f"      Opening URL (direct flights only): {url[:100]}...")
            self.driver.get(url)
            time.sleep(10)
            
            self._close_dialogs()
            
            # Try to apply direct flights filter on results if available
            self._apply_direct_flights_filter()
            
            # Wait for results to load
            time.sleep(5)
            
            # Try to extract price
            price = self._extract_price_from_page([
                "[class*='price']",
                "[class*='fare']",
                "[data-testid*='price']",
                ".price",
                ".fare"
            ])
            
            if price:
                return {
                    'route_code': route['code'],
                    'source': 'Qatar Airways',
                    'source_ar': 'الخطوط القطرية',
                    'source_code': 'AIRL001',
                    'airline': 'Qatar Airways',
                    'price': round(price),
                    'currency': 'QAR',
                    'timestamp': datetime.now().isoformat()
                }
            
            return None
            
        except Exception as e:
            print(f"    Error scraping Qatar Airways: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def scrape_british_airways(self, route: Dict) -> Optional[Dict]:
        """Scrape prices from British Airways"""
        try:
            print(f"    Scraping British Airways for {route['origin']}-{route['destination']}")
            dep_date, ret_date = self._calculate_dates(route['duration_months'])
            
            # Direct URL with search parameters (direct/non-stop flights only)
            url = (f"https://www.britishairways.com/nx/b/airselect/en/usa/book/search?"
                   f"trip=round&arrivalDate={ret_date}&departureDate={dep_date}&"
                   f"from={route['origin_code']}&to={route['destination_code']}&"
                   f"travelClass=economy&adults=1&youngAdults=0&children=0&infants=0&bound=outbound&stops=0")
            
            print(f"      Opening URL (direct flights only): {url[:100]}...")
            self.driver.get(url)
            time.sleep(10)
            
            self._close_dialogs()
            
            # Try to apply direct flights filter on results if available
            self._apply_direct_flights_filter()
            
            # Wait for results
            time.sleep(5)
            
            price = self._extract_price_from_page([
                "[class*='price']",
                "[class*='fare']",
                "[data-testid*='price']",
                ".price",
                ".fare"
            ])
            
            if price:
                return {
                    'route_code': route['code'],
                    'source': 'British Airways',
                    'source_ar': 'الخطوط البريطانية',
                    'source_code': 'AIRL018',
                    'airline': 'British Airways',
                    'price': round(price),
                    'currency': 'QAR',
                    'timestamp': datetime.now().isoformat()
                }
            
            return None
        except Exception as e:
            print(f"    Error scraping British Airways: {e}")
            return None
    
    def scrape_malaysia_airlines(self, route: Dict) -> Optional[Dict]:
        """Scrape prices from Malaysia Airlines"""
        try:
            print(f"    Scraping Malaysia Airlines for {route['origin']}-{route['destination']}")
            dep_date, ret_date = self._calculate_dates(route['duration_months'])
            
            url = "https://www.malaysiaairlines.com/qa/en/home.html"
            print(f"      Opening URL: {url}")
            self.driver.get(url)
            time.sleep(5)
            
            self._close_dialogs()
            
            # Malaysia Airlines requires form filling - would need to implement search form interaction
            # For now, try to extract price if page has results
            price = self._extract_price_from_page([
                "[class*='price']",
                "[class*='fare']",
                "[data-testid*='price']",
                ".price",
                ".fare"
            ])
            
            if price:
                return {
                    'route_code': route['code'],
                    'source': 'Malaysia Airlines',
                    'source_ar': 'الخطوط الماليزية',
                    'source_code': 'AIRL024',
                    'airline': 'Malaysia Airlines',
                    'price': round(price),
                    'currency': 'QAR',
                    'timestamp': datetime.now().isoformat()
                }
            
            return None
        except Exception as e:
            print(f"    Error scraping Malaysia Airlines: {e}")
            return None
    
    def scrape_kuwait_airways(self, route: Dict) -> Optional[Dict]:
        """Scrape prices from Kuwait Airways"""
        try:
            print(f"    Scraping Kuwait Airways for {route['origin']}-{route['destination']}")
            dep_date, ret_date = self._calculate_dates(route['duration_months'])
            
            url = "https://www.kuwaitairways.com/en"
            print(f"      Opening URL: {url}")
            self.driver.get(url)
            time.sleep(5)
            
            self._close_dialogs()
            
            # Kuwait Airways requires form filling - would need to implement search form interaction
            price = self._extract_price_from_page([
                "[class*='price']",
                "[class*='fare']",
                "[data-testid*='price']",
                ".price",
                ".fare"
            ])
            
            if price:
                return {
                    'route_code': route['code'],
                    'source': 'Kuwait Airways',
                    'source_ar': 'الخطوط الكويتية',
                    'source_code': 'AIRL025',
                    'airline': 'Kuwait Airways',
                    'price': round(price),
                    'currency': 'QAR',
                    'timestamp': datetime.now().isoformat()
                }
            
            return None
        except Exception as e:
            print(f"    Error scraping Kuwait Airways: {e}")
            return None
    
    def scrape_turkish_airlines(self, route: Dict) -> Optional[Dict]:
        """Scrape prices from Turkish Airlines"""
        try:
            print(f"    Scraping Turkish Airlines for {route['origin']}-{route['destination']}")
            dep_date, ret_date = self._calculate_dates(route['duration_months'])
            
            # Turkish Airlines uses a different URL structure - need to navigate to booking page first
            url = "https://www.turkishairlines.com/en-qa/flights/booking/availability-international/"
            
            print(f"      Opening URL: {url}")
            self.driver.get(url)
            time.sleep(5)
            
            self._close_dialogs()
            
            # Would need to fill in search form - placeholder for now
            # The URL structure requires session/booking ID which is generated dynamically
            price = self._extract_price_from_page([
                "[class*='price']",
                "[class*='fare']",
                "[data-testid*='price']",
                ".price",
                ".fare"
            ])
            
            if price:
                return {
                    'route_code': route['code'],
                    'source': 'Turkish Airlines',
                    'source_ar': 'الخطوط التركية',
                    'source_code': 'AIRL026',
                    'airline': 'Turkish Airlines',
                    'price': round(price),
                    'currency': 'QAR',
                    'timestamp': datetime.now().isoformat()
                }
            
            return None
        except Exception as e:
            print(f"    Error scraping Turkish Airlines: {e}")
            return None
    
    def scrape_pia(self, route: Dict) -> Optional[Dict]:
        """Scrape prices from Pakistan International Airlines"""
        try:
            print(f"    Scraping PIA for {route['origin']}-{route['destination']}")
            dep_date, ret_date = self._calculate_dates(route['duration_months'])
            
            url = "https://www.piac.com.pk"
            print(f"      Opening URL: {url}")
            self.driver.get(url)
            time.sleep(5)
            
            self._close_dialogs()
            
            # PIA requires form filling - would need to implement search form interaction
            price = self._extract_price_from_page([
                "[class*='price']",
                "[class*='fare']",
                "[data-testid*='price']",
                ".price",
                ".fare"
            ])
            
            if price:
                return {
                    'route_code': route['code'],
                    'source': 'Pakistan International Airlines',
                    'source_ar': 'الخطوط الباكستانية',
                    'source_code': 'AIRL020',
                    'airline': 'PIA',
                    'price': round(price),
                    'currency': 'QAR',
                    'timestamp': datetime.now().isoformat()
                }
            
            return None
        except Exception as e:
            print(f"    Error scraping PIA: {e}")
            return None
    
    def scrape_aggregator(self, source: Dict, route: Dict, airline_name: str = None) -> Optional[Dict]:
        """Scrape prices from travel aggregators (CheapAir, eDreams, KAYAK, ITA Matrix)"""
        try:
            source_name = source['name']
            print(f"    Scraping {source_name} for {route['origin']}-{route['destination']}")
            
            if source_name == 'KAYAK':
                return self._scrape_kayak(route, airline_name)
            elif source_name == 'eDreams':
                return self._scrape_edreams(route, airline_name)
            elif source_name == 'CheapAir':
                return self._scrape_cheapair(route, airline_name)
            elif source_name == 'ITA Matrix':
                return self._scrape_ita_matrix(route, airline_name)
            
            return None
        except Exception as e:
            print(f"    Error scraping {source['name']}: {e}")
            return None
    
    def _scrape_kayak(self, route: Dict, airline_name: str = None) -> Optional[Dict]:
        """Scrape prices from KAYAK"""
        try:
            dep_date, ret_date = self._calculate_dates(route['duration_months'])
            # KAYAK URL format: fs=stops=0 for direct/non-stop flights only
            url = f"https://www.kayak.ae/flights/{route['origin_code']}-{route['destination_code']}/{dep_date}/{ret_date}?ucs=bzx8kr&sort=bestflight_a&fs=stops=0"
            
            print(f"      Opening URL (direct flights only): {url[:100]}...")
            self.driver.get(url)
            time.sleep(10)
            
            self._close_dialogs()
            
            # Try to apply "Direct flights only" filter if not in URL
            self._apply_direct_flights_filter()
            
            # Wait for results to load - KAYAK may take time
            print(f"      Waiting for results to load...")
            time.sleep(10)
            
            # Try to close any popups or overlays
            try:
                close_buttons = self.driver.find_elements(By.CSS_SELECTOR, 
                    "button[aria-label*='Close'], button[aria-label*='Dismiss'], .close-button, [class*='close']"
                )
                for btn in close_buttons[:3]:  # Try first 3 close buttons
                    try:
                        if btn.is_displayed():
                            btn.click()
                            time.sleep(1)
                    except:
                        continue
            except:
                pass
            
            # Scroll to trigger lazy loading
            self.driver.execute_script("window.scrollTo(0, 500);")
            time.sleep(3)
            
            # Use the helper method to extract price
            price = self._extract_price_from_page([
                "[data-test-id='price']",
                "[data-testid='price']",
                ".price-text",
                ".Flights-Price-FlightPrice",
                "[class*='price']",
                "[class*='Price']",
                ".result-price",
                "[data-test-id='result-price']",
                "span[class*='price']",
                "div[class*='price']"
            ])
            
            if price:
                return {
                    'route_code': route['code'],
                    'source': 'KAYAK',
                    'source_ar': 'Kayak',
                    'source_code': 'AIRL028',
                    'airline': airline_name or 'Various',
                    'price': round(price),
                    'currency': 'QAR',
                    'timestamp': datetime.now().isoformat()
                }
            
            print(f"      ✗ Could not find price on page")
            return None
        except Exception as e:
            print(f"    Error in KAYAK scraping: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def _scrape_edreams(self, route: Dict, airline_name: str = None) -> Optional[Dict]:
        """Scrape prices from eDreams"""
        try:
            dep_date, ret_date = self._calculate_dates(route['duration_months'])
            
            # eDreams uses hash-based routing; directOnly=1 for direct flights only
            url = (f"https://www.edreams.qa/travel/#results/"
                   f"type=R;from={route['origin_code']};to={route['destination_code']};"
                   f"dep={dep_date};ret={ret_date};"
                   f"buyPath=FLIGHTS_HOME_SEARCH_FORM;internalSearch=true;directOnly=true")
            
            print(f"      Opening URL (direct flights only): {url[:100]}...")
            self.driver.get(url)
            time.sleep(10)
            
            self._close_dialogs()
            
            # Try to click "Direct flights only" if visible (رحلات مباشرة فقط)
            self._apply_direct_flights_filter()
            
            # Wait for results to load (eDreams may take time)
            time.sleep(8)
            
            # Scroll to trigger loading
            self.driver.execute_script("window.scrollTo(0, 500);")
            time.sleep(3)
            
            price = self._extract_price_from_page([
                "[class*='price']",
                "[class*='fare']",
                "[data-testid*='price']",
                ".price",
                ".fare",
                "[class*='Price']"
            ])
            
            if price:
                return {
                    'route_code': route['code'],
                    'source': 'eDreams',
                    'source_ar': 'edreams',
                    'source_code': 'AIRL030',
                    'airline': airline_name or 'Various',
                    'price': round(price),
                    'currency': 'QAR',
                    'timestamp': datetime.now().isoformat()
                }
            
            return None
        except Exception as e:
            print(f"    Error in eDreams scraping: {e}")
            return None
    
    def _scrape_cheapair(self, route: Dict, airline_name: str = None) -> Optional[Dict]:
        """Scrape prices from CheapAir (CheapoAir)"""
        try:
            dep_date, ret_date = self._calculate_dates(route['duration_months'])
            
            # Note: The URL shows cheapoair.com, not cheapair.com
            # Format dates as MM/DD/YYYY for CheapAir
            dep_date_formatted = datetime.strptime(dep_date, '%Y-%m-%d').strftime('%m/%d/%Y')
            ret_date_formatted = datetime.strptime(ret_date, '%Y-%m-%d').strftime('%m/%d/%Y')
            
            # nonstop=1 for direct flights only
            url = (f"https://www.cheapoair.com/air/listing?"
                   f"&d1={route['origin_code']}&r1={route['destination_code']}&"
                   f"dt1={dep_date_formatted}&dtype1=A&rtype1=C&"
                   f"d2={route['destination_code']}&r2={route['origin_code']}&"
                   f"dt2={ret_date_formatted}&dtype2=C&rtype2=A&"
                   f"tripType=ROUNDTRIP&cl=ECONOMY&ad=1&se=0&ch=0&infs=0&infl=0&nonstop=1")
            
            print(f"      Opening URL (direct flights only): {url[:100]}...")
            self.driver.get(url)
            time.sleep(10)
            
            self._close_dialogs()
            
            time.sleep(5)
            
            price = self._extract_price_from_page([
                "[class*='price']",
                "[class*='fare']",
                "[data-testid*='price']",
                ".price",
                ".fare"
            ])
            
            if price:
                return {
                    'route_code': route['code'],
                    'source': 'CheapAir',
                    'source_ar': 'cheapair',
                    'source_code': 'AIRL028',
                    'airline': airline_name or 'Various',
                    'price': round(price),
                    'currency': 'QAR',
                    'timestamp': datetime.now().isoformat()
                }
            
            return None
        except Exception as e:
            print(f"    Error in CheapAir scraping: {e}")
            return None
    
    def _scrape_ita_matrix(self, route: Dict, airline_name: str = None) -> Optional[Dict]:
        """Scrape prices from ITA Matrix"""
        try:
            dep_date, ret_date = self._calculate_dates(route['duration_months'])
            
            # ITA Matrix uses encoded search parameters
            # The search parameter is base64 encoded JSON
            import json as json_lib
            
            search_params = {
                "type": "round-trip",
                "slices": [{
                    "origin": [route['origin_code']],
                    "dest": [route['destination_code']],
                    "dates": {
                        "searchDateType": "specific",
                        "departureDate": dep_date,
                        "departureDateType": "depart",
                        "departureDateModifier": "0",
                        "departureDatePreferredTimes": [],
                        "returnDate": ret_date,
                        "returnDateType": "depart",
                        "returnDateModifier": "0",
                        "returnDatePreferredTimes": []
                    }
                }],
                "options": {
                    "cabin": "COACH",
                    "stops": "-1",   # -1 means non-stop only
                    "extraStops": "0",  # 0 = no extra stops (direct only)
                    "allowAirportChanges": "true",
                    "showOnlyAvailable": "true"
                },
                "pax": {
                    "adults": "1"
                }
            }
            
            # Encode the search parameters
            search_json = json_lib.dumps(search_params)
            search_encoded = base64.b64encode(search_json.encode()).decode()
            
            url = f"https://matrix.itasoftware.com/flights?search={search_encoded}"
            
            print(f"      Opening ITA Matrix URL...")
            self.driver.get(url)
            time.sleep(10)
            
            self._close_dialogs()
            
            time.sleep(5)
            
            price = self._extract_price_from_page([
                "[class*='price']",
                "[class*='fare']",
                "[data-testid*='price']",
                ".price",
                ".fare"
            ])
            
            if price:
                return {
                    'route_code': route['code'],
                    'source': 'ITA Matrix',
                    'source_ar': 'matrix',
                    'source_code': 'AIRL028',
                    'airline': airline_name or 'Various',
                    'price': round(price),
                    'currency': 'QAR',
                    'timestamp': datetime.now().isoformat()
                }
            
            return None
        except Exception as e:
            print(f"    Error in ITA Matrix scraping: {e}")
            return None
    
    def _apply_direct_flights_filter(self):
        """Try to apply 'Direct flights only' / 'Nonstop' filter on the page"""
        try:
            # Common labels for direct/non-stop filter (English and Arabic)
            direct_keywords = [
                'direct', 'nonstop', 'non-stop', 'non stop',
                'رحلات مباشرة', 'مباشرة فقط', 'بدون توقف'
            ]
            # Selectors for checkboxes, filters, buttons
            selectors = [
                "[data-testid*='nonstop']",
                "[data-testid*='direct']",
                "[data-test-id*='nonstop']",
                "[data-test-id*='stops-0']",
                "input[type='checkbox'][id*='direct']",
                "input[type='checkbox'][id*='nonstop']",
                "label:has(input[type='checkbox'])",
                "[aria-label*='Nonstop']",
                "[aria-label*='Direct']",
                "button[class*='stops']",
                "[class*='nonstop']",
                "[class*='direct-only']",
                "a[href*='stops=0']",
                "a[href*='nonstop']"
            ]
            for selector in selectors:
                try:
                    elems = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for elem in elems:
                        try:
                            text = (elem.text or elem.get_attribute('aria-label') or '').lower()
                            if any(kw in text for kw in ['direct', 'nonstop', 'non-stop', 'non stop']):
                                if elem.is_displayed():
                                    elem.click()
                                    time.sleep(2)
                                    print(f"      Applied direct flights filter")
                                    return
                        except:
                            continue
                    if elems and elems[0].is_displayed():
                        elems[0].click()
                        time.sleep(2)
                        print(f"      Applied direct flights filter")
                        return
                except:
                    continue
            # Try by link text / partial text
            try:
                for kw in ['Nonstop', 'Direct', 'Non-stop', 'رحلات مباشرة']:
                    try:
                        link = self.driver.find_element(By.LINK_TEXT, kw)
                        if link.is_displayed():
                            link.click()
                            time.sleep(2)
                            print(f"      Applied direct flights filter")
                            return
                    except:
                        try:
                            link = self.driver.find_element(By.PARTIAL_LINK_TEXT, kw)
                            if link.is_displayed():
                                link.click()
                                time.sleep(2)
                                print(f"      Applied direct flights filter")
                                return
                        except:
                            continue
            except:
                pass
        except Exception as e:
            pass  # Silently ignore if filter not found
    
    def _close_dialogs(self):
        """Close cookie/consent dialogs"""
        try:
            consent_selectors = [
                "button#onetrust-accept-btn-handler",
                "button[id*='accept']",
                "button[class*='accept']",
                "[data-testid='cookie-consent-accept']",
                "button[aria-label*='Accept']",
                "button[aria-label*='Close']"
            ]
            for selector in consent_selectors:
                try:
                    consent_btn = WebDriverWait(self.driver, 3).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                    )
                    consent_btn.click()
                    time.sleep(1)
                    break
                except:
                    continue
        except:
            pass
    
    def scrape_all(self) -> Dict:
        """Scrape prices for all routes from all sources"""
        print("\n" + "="*60)
        print("FLIGHT PRICE SCRAPER")
        print("="*60)
        print(f"Scraping {len(self.routes)} routes from {len(self.sources)} sources...")
        print("="*60 + "\n")
        
        self._setup_driver(headless=self.headless)
        
        results = {
            'timestamp': datetime.now().isoformat(),
            'routes': []
        }
        
        for route in self.routes:
            print(f"\n[{route['code']}] Route: {route['origin']} - {route['destination']}")
            route_results = {
                'route': route,
                'prices': []
            }
            
            for source in self.sources:
                print(f"  Source: {source['name']} ({source['type']})")
                
                try:
                    if source['type'] == 'airline':
                        # Scrape directly from airline
                        price_data = None
                        if source['name'] == 'Qatar Airways':
                            price_data = self.scrape_qatar_airways(route)
                        elif source['name'] == 'British Airways':
                            price_data = self.scrape_british_airways(route)
                        elif source['name'] == 'Malaysia Airlines':
                            price_data = self.scrape_malaysia_airlines(route)
                        elif source['name'] == 'Kuwait Airways':
                            price_data = self.scrape_kuwait_airways(route)
                        elif source['name'] == 'Turkish Airlines':
                            price_data = self.scrape_turkish_airlines(route)
                        elif source['name'] == 'Pakistan International Airlines':
                            price_data = self.scrape_pia(route)
                        
                        if price_data:
                            price_data['route_code'] = route['code']
                            price_data['source'] = source['name']
                            price_data['source_ar'] = source['name_ar']
                            price_data['source_code'] = source['source_code']
                            route_results['prices'].append(price_data)
                            print(f"    ✓ Price found: {price_data.get('price')} {price_data.get('currency', '')}")
                        else:
                            print(f"    ✗ No price found")
                    
                    elif source['type'] == 'aggregator':
                        # For aggregators, we can search for specific airlines
                        # For now, search without airline filter
                        price_data = self.scrape_aggregator(source, route)
                        
                        if price_data:
                            price_data['route_code'] = route['code']
                            price_data['source'] = source['name']
                            price_data['source_ar'] = source['name_ar']
                            price_data['source_code'] = source['source_code']
                            route_results['prices'].append(price_data)
                            print(f"    ✓ Price found: {price_data.get('price')} {price_data.get('currency', '')}")
                        else:
                            print(f"    ✗ No price found")
                    
                    # Delay between requests
                    time.sleep(3)
                    
                except Exception as e:
                    print(f"    ✗ Error: {e}")
                    continue
            
            results['routes'].append(route_results)
            
            # Write this route to Excel immediately so data is saved as we go
            try:
                self.append_route_to_excel(route_results, self.excel_path)
                print(f"  ✓ Saved route {route['code']} to Excel")
            except Exception as e:
                print(f"  ✗ Could not save to Excel: {e}")
        
        self._close_driver()
        
        # Print summary
        total_prices = sum(len(r['prices']) for r in results['routes'])
        print("\n" + "="*60)
        print(f"Scraping completed: {total_prices} prices found")
        print("="*60)
        
        return results
    
    def _apply_border(self, cell):
        """Apply thin border to a cell"""
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        cell.border = thin_border
    
    def _flight_header_row(self, ws) -> Optional[int]:
        """Return the 1-based row number where flight price headers are (column 6 contains FLIGHT_HEADER_MARKER), or None."""
        for r in range(1, ws.max_row + 1):
            if FLIGHT_HEADER_MARKER in str(ws.cell(row=r, column=6).value or ''):
                return r
        return None

    def _ensure_route_block(self, ws) -> int:
        """Ensure the sheet has route block at top (row 1 = headers, rows 2+ = data). Returns the row after the route block (0 if flight-only sheet)."""
        col3 = str(ws.cell(row=1, column=3).value or '')
        if 'Class' in col3 or 'الدرجة' in col3:
            return 0
        header_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        if ws.cell(row=1, column=1).value != 'Code':
            for col, h in enumerate(ROUTE_HEADERS, 1):
                c = ws.cell(row=1, column=col)
                c.value = h
                c.font = Font(bold=True)
                c.fill = header_fill
        if not ws.cell(row=2, column=1).value:
            for row_idx, r in enumerate(self._get_default_routes(), 2):
                ws.cell(row=row_idx, column=1).value = r['code']
                ws.cell(row=row_idx, column=2).value = r['commodity_ar']
                ws.cell(row=row_idx, column=3).value = r['origin']
                ws.cell(row=row_idx, column=4).value = r['origin_code']
                ws.cell(row=row_idx, column=5).value = r['destination']
                ws.cell(row=row_idx, column=6).value = r['destination_code']
                ws.cell(row=row_idx, column=7).value = r['duration_months']
            print(f"  Created route block at top of sheet with default routes (edit in Excel to add/remove routes)")
        last_route_row = 1
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=1).value and FLIGHT_HEADER_MARKER not in str(ws.cell(row=r, column=6).value or ''):
                last_route_row = r
            else:
                break
        return last_route_row + 1
    
    def _prepare_excel_for_export(self, filename: str):
        """Load or create workbook; single sheet with route block at top and flight prices below. Return (wb, ws, date_col, next_row, thin_border, avg_fill)."""
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        header_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        flight_headers = [
            'Code', 'Commodity', 'الدرجة المقابلة لها في الخطوط (Class equivalent in airlines)',
            'CPI-Flag', 'رمز المصدر (Source Code)', 'وكالات الخطوط (Flight Agencies)'
        ]
        file_exists = os.path.exists(filename)
        if file_exists:
            wb = load_workbook(filename)
            if FLIGHT_PRICES_SHEET_NAME in wb.sheetnames:
                ws = wb[FLIGHT_PRICES_SHEET_NAME]
            else:
                ws = wb.active
                ws.title = FLIGHT_PRICES_SHEET_NAME
            ws.sheet_view.rightToLeft = True
            row_after_routes = self._ensure_route_block(ws)
            hrow = self._flight_header_row(ws)
            if hrow is None:
                flight_header_row = row_after_routes + 1
                for col_idx, header in enumerate(flight_headers, 1):
                    cell = ws.cell(row=flight_header_row, column=col_idx)
                    cell.value = header
                    cell.font = Font(bold=True)
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = thin_border
                max_col = 6
            else:
                flight_header_row = hrow
                max_col = ws.max_column
                if max_col < 7:
                    max_col = 7
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = FLIGHT_PRICES_SHEET_NAME
            ws.sheet_view.rightToLeft = True
            row_after_routes = self._ensure_route_block(ws)
            flight_header_row = row_after_routes + 1
            for col_idx, header in enumerate(flight_headers, 1):
                cell = ws.cell(row=flight_header_row, column=col_idx)
                cell.value = header
                cell.font = Font(bold=True)
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border
            max_col = 6
        
        scheduled_dates = self._get_scheduled_dates_through_2026()

        def _norm_date_header(val):
            if not val:
                return None
            s = str(val).strip()
            try:
                d = datetime.strptime(s, '%d-%b')
                return d.strftime('%d-%b')
            except ValueError:
                try:
                    d = datetime.strptime(s, '%d-%b-%Y')
                    return d.strftime('%d-%b')
                except ValueError:
                    return s
        existing_headers = {}
        for col in range(7, max_col + 1):
            val = ws.cell(row=flight_header_row, column=col).value
            key = _norm_date_header(val)
            if key:
                existing_headers[key] = col
        next_col = max_col + 1
        for header_text, _ in scheduled_dates:
            if header_text in existing_headers:
                continue
            cell = ws.cell(row=flight_header_row, column=next_col)
            cell.value = header_text
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
            existing_headers[header_text] = next_col
            next_col += 1
        max_col = next_col - 1
        date_text = self._get_current_run_date_header()
        date_col = existing_headers.get(date_text)
        if date_col is None:
            date_col = max_col + 1
            cell = ws.cell(row=flight_header_row, column=date_col)
            cell.value = date_text
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
            max_col = date_col
        for col in range(1, max_col + 1):
            c = ws.cell(row=flight_header_row, column=col)
            c.fill = header_fill
            c.font = Font(bold=True)
            c.border = thin_border
            c.alignment = Alignment(horizontal='center', vertical='center')
        row = flight_header_row + 1
        while ws.cell(row=row, column=1).value is not None:
            row += 1
        avg_fill = PatternFill(start_color='FFE4B5', end_color='FFE4B5', fill_type='solid')
        return (wb, ws, date_col, row, thin_border, avg_fill)
    
    def _write_route_to_sheet(self, ws, route_result: Dict, row: int, date_col: int, thin_border, avg_fill) -> int:
        """Write one route's data (prices + averages) to ws starting at row. Returns next row after this route."""
        route = route_result['route']
        prices = route_result.get('prices', [])
        route_start_row = row
        sorted_prices = sorted(prices, key=lambda x: (x.get('airline', 'Various'), x.get('source', '')))
        for price_data in sorted_prices:
            route = route_result['route']
            ws.cell(row=row, column=1).value = route['code']
            ws.cell(row=row, column=1).border = thin_border
            ws.cell(row=row, column=1).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=row, column=2).value = route['commodity_ar']
            ws.cell(row=row, column=2).border = thin_border
            ws.cell(row=row, column=2).alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
            ws.cell(row=row, column=3).value = 'Economy'
            ws.cell(row=row, column=3).border = thin_border
            ws.cell(row=row, column=3).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=row, column=4).value = 'Y'
            ws.cell(row=row, column=4).border = thin_border
            ws.cell(row=row, column=4).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=row, column=5).value = price_data.get('source_code', '')
            ws.cell(row=row, column=5).border = thin_border
            ws.cell(row=row, column=5).alignment = Alignment(horizontal='center', vertical='center')
            agency_name = price_data.get('source_ar', price_data.get('source', ''))
            airline_name = price_data.get('airline', '')
            source_name = price_data.get('source', '')
            if airline_name and airline_name != 'Various':
                if source_name in ['KAYAK', 'eDreams', 'CheapAir', 'ITA Matrix']:
                    val = f"Kayak عبر {airline_name}" if source_name == 'KAYAK' else (f"matrix عبر {airline_name}" if source_name == 'ITA Matrix' else f"{agency_name} عبر {airline_name}")
                else:
                    val = agency_name
            else:
                val = agency_name
            ws.cell(row=row, column=6).value = val
            ws.cell(row=row, column=6).border = thin_border
            ws.cell(row=row, column=6).alignment = Alignment(horizontal='right', vertical='center')
            p = price_data.get('price')
            if p:
                ws.cell(row=row, column=date_col).value = p
                ws.cell(row=row, column=date_col).number_format = '0'
            ws.cell(row=row, column=date_col).border = thin_border
            ws.cell(row=row, column=date_col).alignment = Alignment(horizontal='center', vertical='center')
            row += 1
        airline_avg_groups = {}
        for price_data in prices:
            airline = price_data.get('airline', 'Various')
            if airline != 'Various':
                if airline not in airline_avg_groups:
                    airline_avg_groups[airline] = []
                if price_data.get('price') and price_data.get('price') > 0:
                    airline_avg_groups[airline].append(price_data.get('price'))
        airline_ar_map = {'Qatar Airways': 'القطرية', 'British Airways': 'البريطانية', 'Malaysia Airlines': 'الماليزية', 'Kuwait Airways': 'الكويتية', 'Turkish Airlines': 'التركية', 'Pakistan International Airlines': 'الباكستانية', 'PIA': 'الباكستانية'}
        for airline, price_list in airline_avg_groups.items():
            if len(price_list) <= 1:
                continue
            avg_price = sum(price_list) / len(price_list)
            for col in range(1, 7):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border
                cell.fill = avg_fill
                cell.alignment = Alignment(horizontal='center' if col != 2 and col != 6 else 'right', vertical='center')
            ws.cell(row=row, column=1).value = route['code']
            ws.cell(row=row, column=2).value = route['commodity_ar']
            ws.cell(row=row, column=3).value = 'Economy'
            ws.cell(row=row, column=4).value = 'N-averages'
            ws.cell(row=row, column=5).value = ''
            ws.cell(row=row, column=6).value = f"متوسط المصادر للخطوط {airline_ar_map.get(airline, airline)}"
            pc = ws.cell(row=row, column=date_col)
            pc.value = round(avg_price)
            pc.number_format = '0'
            pc.font = Font(bold=True)
            pc.border = thin_border
            pc.fill = avg_fill
            pc.alignment = Alignment(horizontal='center', vertical='center')
            for c in range(7, date_col):
                fc = ws.cell(row=row, column=c)
                fc.fill = avg_fill
                fc.border = thin_border
            row += 1
        all_valid_prices = [p.get('price') for p in prices if p.get('price') and p.get('price') > 0]
        if all_valid_prices:
            overall_avg = sum(all_valid_prices) / len(all_valid_prices)
            for col in range(1, 7):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border
                cell.fill = avg_fill
                cell.alignment = Alignment(horizontal='center' if col not in (2, 6) else 'right', vertical='center')
            ws.cell(row=row, column=1).value = route['code']
            ws.cell(row=row, column=2).value = route['commodity_ar']
            ws.cell(row=row, column=3).value = 'Economy'
            ws.cell(row=row, column=4).value = 'Y-averages'
            ws.cell(row=row, column=5).value = ''
            ws.cell(row=row, column=6).value = "متوسط المصادر"
            pc = ws.cell(row=row, column=date_col)
            pc.value = round(overall_avg)
            pc.number_format = '0'
            pc.font = Font(bold=True)
            pc.border = thin_border
            pc.fill = avg_fill
            pc.alignment = Alignment(horizontal='center', vertical='center')
            for c in range(7, date_col):
                fc = ws.cell(row=row, column=c)
                fc.fill = avg_fill
                fc.border = thin_border
            row += 1
        route_end_row = row - 1
        if route_end_row >= route_start_row and route_end_row > route_start_row:
            try:
                for merged_range in list(ws.merged_cells.ranges):
                    if merged_range.min_col == 2 and merged_range.max_col == 2 and merged_range.min_row <= route_start_row <= merged_range.max_row:
                        ws.unmerge_cells(str(merged_range))
                ws.merge_cells(f'B{route_start_row}:B{route_end_row}')
                ws.cell(row=route_start_row, column=2).alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
            except Exception:
                pass
        return row
    
    def append_route_to_excel(self, route_result: Dict, filename: str = None):
        """Append one route's data to the Excel file immediately (called after each route is scraped)."""
        filename = filename or self.excel_path
        wb, ws, date_col, row, thin_border, avg_fill = self._prepare_excel_for_export(filename)
        row = self._write_route_to_sheet(ws, route_result, row, date_col, thin_border, avg_fill)
        wb.save(filename)
    
    def export_to_excel(self, results: Dict, filename: str = None):
        """Export all flight prices to Excel. If data was already written per-route, this can re-write or just save column widths."""
        filename = filename or self.excel_path
        print(f"\n{'='*60}")
        print("EXPORTING TO EXCEL")
        print(f"{'='*60}")
        
        if 'error' in results:
            print(f"Error: Cannot export - {results['error']}")
            return False
        
        wb, ws, date_col, row, thin_border, avg_fill = self._prepare_excel_for_export(filename)
        initial_row = row
        for route_result in results.get('routes', []):
            row = self._write_route_to_sheet(ws, route_result, row, date_col, thin_border, avg_fill)
        
        # Auto-adjust column widths
        ws.column_dimensions['A'].width = 12  # Code
        ws.column_dimensions['B'].width = 60  # Commodity
        ws.column_dimensions['C'].width = 25  # Class
        ws.column_dimensions['D'].width = 12  # CPI-Flag
        ws.column_dimensions['E'].width = 15  # Source Code
        ws.column_dimensions['F'].width = 35  # Flight Agencies
        
        for col in range(7, date_col + 1):
            ws.column_dimensions[get_column_letter(col)].width = 15
        
        ws.sheet_view.rightToLeft = True
        
        date_text = self._get_current_run_date_header()
        try:
            wb.save(filename)
            rows_written = row - initial_row
            print(f"\n✓ Prices exported to {filename} (RTL layout)")
            print(f"  Rows written: {rows_written}")
            print(f"  Date column: {date_text} (column {date_col})")
            return True
        except Exception as e:
            print(f"\n✗ Error saving Excel file: {e}")
            import traceback
            traceback.print_exc()
            return False


def main():
    """Main function"""
    import sys
    
    # Set headless=False to see the browser, True to run in background
    scraper = FlightPriceScraper(headless=False)
    
    results = scraper.scrape_all()
    
    # Save to JSON
    json_file = 'flight_prices.json'
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nPrices saved to {json_file}")
    
    # Export to Excel (uses scraper.excel_path; routes are read from sheet "Routes")
    scraper.export_to_excel(results)


if __name__ == "__main__":
    main()
