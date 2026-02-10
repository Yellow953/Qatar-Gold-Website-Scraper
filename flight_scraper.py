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

6. **Real prices vs skipped:**
   - Real search URLs (so we can get real fares): Qatar Airways, British Airways, CheapAir, eDreams, KAYAK, ITA Matrix.
   - Malaysia, Kuwait, Turkish, PIA: no direct search URL; we skip them (return no price) to avoid fake numbers from homepages.
   - Min/max QAR caps filter out wrong elements (one-way, multi-pax, wrong currency).
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
        path = os.path.abspath(self.excel_path)
        if not os.path.exists(path):
            print(f"  Excel file not found: {path}")
            return None
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            if FLIGHT_PRICES_SHEET_NAME in wb.sheetnames:
                ws = wb[FLIGHT_PRICES_SHEET_NAME]
            else:
                ws = wb.active
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
                print(f"  Reading routes from Excel: {path}")
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
                'commodity_ar': 'كلفة تذكرة دوحة_ كراتشي _ دوحة لمدة 6 اشهر ( التذكرة سياحية semi flexble)',
                'commodity_en': 'Cost of a Doha - Karachi - Doha ticket for 6 (semi flexible tourist ticket) months',
                'ticket_type': 'Semi flexible',
                'duration_months': 6
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
                'commodity_ar': 'كلفة تذكرة دوحة_ بومباي _ دوحة لمدة 6 اشهر ( التذكرة سياحية semi flexble)',
                'commodity_en': 'Cost of a Doha - Mumbai - Doha ticket for 6 (semi flexible tourist ticket) months',
                'ticket_type': 'Semi flexible',
                'duration_months': 6
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
        """Return the date header for this run: use today's date so every run gets the next column (even if not 4/10/17/24)."""
        return datetime.now().strftime('%d-%b')
    
    # Approximate exchange rates to QAR (Qatari Riyal) - update periodically if needed
    CURRENCY_TO_QAR = {
        'QAR': 1.0, 'QR': 1.0, 'ر.ق': 1.0,
        'USD': 3.64, 'US$': 3.64, '$': 3.64,
        'AED': 1.0, 'د.إ': 1.0, 'DH': 1.0,
        'EUR': 3.90, '€': 3.90,
        'GBP': 4.60, '£': 4.60,
        'SAR': 0.97, 'SR': 0.97,
        'BHD': 9.65, 'KWD': 11.85, 'OMR': 9.46,
    }
    
    def _detect_currency_from_text(self, text: str) -> Tuple[Optional[float], str]:
        """Extract numeric amount and detect currency from text (e.g. '1,049 USD', '$1,049', '3,594 AED').
        Returns (amount, currency_code). currency_code is 'QAR' if unknown (assume already QAR)."""
        if not text or not text.strip():
            return None, 'QAR'
        text_upper = text.upper().strip()
        amount = None
        detected = 'QAR'
        # Currency patterns: symbol/code before or after number
        patterns_currency = [
            (r'(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)\s*(QAR|QR|USD|US\$|\$|AED|EUR|€|GBP|£|SAR|BHD|KWD|OMR)', 1, 2),
            (r'(QAR|QR|USD|US\$|\$|AED|EUR|€|GBP|£|SAR|BHD|KWD|OMR)\s*(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)', 2, 1),
            (r'\$\s*(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)', 1, 'USD'),
            (r'(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)\s*\$', 1, 'USD'),
            (r'(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)\s*(?:AED|د\.إ)', 1, 'AED'),
            (r'(?:AED|د\.إ)\s*(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)', 1, 'AED'),
        ]
        for pat, num_grp, cur_grp in patterns_currency:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                try:
                    num_str = m.group(num_grp).replace(',', '')
                    amount = float(num_str)
                    cur = m.group(cur_grp) if isinstance(cur_grp, int) else cur_grp
                    if cur in ('$', 'US$'):
                        detected = 'USD'
                    elif cur in ('€',):
                        detected = 'EUR'
                    elif cur in ('£',):
                        detected = 'GBP'
                    else:
                        detected = (cur or 'QAR').strip().upper()
                    if amount and amount > 0:
                        return amount, detected
                except (IndexError, ValueError, TypeError):
                    pass
        cleaned = re.sub(r'[^\d.,]', '', text)
        cleaned = cleaned.replace(',', '')
        try:
            amount = float(cleaned)
            if amount > 0:
                return amount, 'QAR'
        except ValueError:
            pass
        return None, 'QAR'
    
    def _convert_to_qar(self, amount: float, currency_code: str) -> float:
        """Convert amount from given currency to QAR. Returns amount in QAR."""
        if amount is None or amount <= 0:
            return 0.0
        code = (currency_code or 'QAR').strip().upper()
        if code in ('$', 'US$'):
            code = 'USD'
        rate = self.CURRENCY_TO_QAR.get(code, 1.0)
        return round(amount * rate, 0)
    
    def _extract_price_from_text(self, text: str) -> Optional[float]:
        """Extract numeric price from text (legacy; use _detect_currency_from_text for currency)."""
        am, _ = self._detect_currency_from_text(text)
        return am
    
    # Long-haul destinations: use 1000 QAR min to filter one-way/wrong numbers but allow low fares. Short-haul stays 500.
    LONG_HAUL_DESTINATION_CODES = {'LHR', 'LON', 'JFK', 'NYC', 'IST', 'BKK', 'KUL', 'TBS', 'FCO', 'CDG', 'MAD', 'FRA', 'MUC', 'AMS', 'SYD', 'MEL', 'SIN', 'HKG', 'NRT', 'TYO'}
    
    def _min_price_for_route(self, route: Optional[Dict]) -> int:
        """Minimum plausible round-trip economy fare in QAR (avoids one-way or wrong numbers)."""
        if not route:
            return 500
        dest = (route.get('destination_code') or '').upper()
        if dest in self.LONG_HAUL_DESTINATION_CODES:
            return 1000
        return 500

    # Plausible max round-trip economy in QAR (filters wrong page elements / totals for multiple pax).
    MAX_QAR_LONG_HAUL = 15000
    MAX_QAR_SHORT_HAUL = 8000

    def _max_price_for_route(self, route: Optional[Dict]) -> int:
        """Maximum plausible round-trip economy fare in QAR (rejects fake/wrong numbers from wrong pages)."""
        if not route:
            return 50000
        dest = (route.get('destination_code') or '').upper()
        if dest in self.LONG_HAUL_DESTINATION_CODES:
            return self.MAX_QAR_LONG_HAUL
        return self.MAX_QAR_SHORT_HAUL

    def _extract_price_from_page(self, selectors: List[str], min_price: int = None, max_price: int = None, route: Dict = None) -> Optional[float]:
        """Extract price from page; detect currency (USD, AED, etc.) and convert to QAR.
        Prefers round-trip/total price when multiple candidates. Uses route-aware min/max to avoid wrong fares. Returns price in QAR or None."""
        if min_price is None:
            min_price = self._min_price_for_route(route)
        if max_price is None and route is not None:
            max_price = self._max_price_for_route(route)
        if max_price is None:
            max_price = 50000
        candidates_with_currency = []
        
        for selector in selectors:
            try:
                price_elems = self.driver.find_elements(By.CSS_SELECTOR, selector)
                for price_elem in price_elems[:15]:
                    try:
                        price_text = (price_elem.text or '').strip()
                        if not price_text:
                            continue
                        amount, currency = self._detect_currency_from_text(price_text)
                        if not amount or amount <= 0:
                            continue
                        qar = self._convert_to_qar(amount, currency)
                        if min_price <= qar <= max_price:
                            is_total = any(k in price_text.lower() for k in ('total', 'round', 'return', 'round-trip', 'roundtrip', 'رحلة ذهاب وعودة'))
                            candidates_with_currency.append((qar, is_total, amount, currency))
                    except Exception:
                        continue
                if candidates_with_currency:
                    break
            except Exception:
                continue
        
        if candidates_with_currency:
            candidates_with_currency.sort(key=lambda x: (not x[1], x[0]))
            best = candidates_with_currency[0]
            qar_val = int(round(best[0]))
            if best[3] != 'QAR':
                print(f"      ✓ Found price: {best[2]} {best[3]} → {qar_val} QAR")
            else:
                print(f"      ✓ Found price: {qar_val} QAR")
            return float(qar_val)
        
        page_text = ""
        try:
            page_text = self.driver.page_source
        except Exception:
            pass
        
        currency_patterns = [
            (r'(?:QAR|QR)\s*(\d{1,3}(?:,\d{3})*)', 'QAR'),
            (r'\$\s*(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)', 'USD'),
            (r'(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)\s*USD', 'USD'),
            (r'(?:AED|د\.إ)\s*(\d{1,3}(?:,\d{3})*)', 'AED'),
            (r'(\d{1,3}(?:,\d{3})*)\s*AED', 'AED'),
            (r'(\d{4,6})', 'QAR'),
            (r'(\d{1,3}(?:,\d{3})+(?:\.\d{2})?)', 'QAR'),
        ]
        all_candidates = []
        for pattern, cur in currency_patterns:
            matches = re.findall(pattern, page_text, re.IGNORECASE)
            for match in list(matches)[:25]:
                try:
                    val = float(str(match).replace(',', ''))
                    qar = self._convert_to_qar(val, cur)
                    if min_price <= qar <= max_price:
                        all_candidates.append((qar, val, cur))
                except (ValueError, TypeError):
                    continue
        if all_candidates:
            all_candidates.sort(key=lambda x: x[0], reverse=True)
            qar, orig, cur = all_candidates[0]
            qar_int = int(round(qar))
            if cur != 'QAR':
                print(f"      ✓ Extracted from page: {orig} {cur} → {qar_int} QAR (round-trip total)")
            else:
                print(f"      ✓ Extracted price from page source: {qar_int} QAR")
            return float(qar_int)
        
        return None
    
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
        """Scrape prices from Malaysia Airlines. No direct search URL – would need to automate search form; skipping to avoid fake prices."""
        try:
            print(f"    Scraping Malaysia Airlines for {route['origin']}-{route['destination']}")
            print(f"      (Search not implemented – opening homepage would give wrong numbers; skipping)")
            return None
        except Exception as e:
            print(f"    Error scraping Malaysia Airlines: {e}")
            return None
    
    def scrape_kuwait_airways(self, route: Dict) -> Optional[Dict]:
        """Scrape prices from Kuwait Airways. No direct search URL – would need to automate search form; skipping to avoid fake prices."""
        try:
            print(f"    Scraping Kuwait Airways for {route['origin']}-{route['destination']}")
            print(f"      (Search not implemented – opening homepage would give wrong numbers; skipping)")
            return None
        except Exception as e:
            print(f"    Error scraping Kuwait Airways: {e}")
            return None
    
    def scrape_turkish_airlines(self, route: Dict) -> Optional[Dict]:
        """Scrape prices from Turkish Airlines. No direct search URL – would need to automate search form; skipping to avoid fake prices."""
        try:
            print(f"    Scraping Turkish Airlines for {route['origin']}-{route['destination']}")
            print(f"      (Search not implemented – generic booking page would give wrong numbers; skipping)")
            return None
        except Exception as e:
            print(f"    Error scraping Turkish Airlines: {e}")
            return None
    
    def scrape_pia(self, route: Dict) -> Optional[Dict]:
        """Scrape prices from Pakistan International Airlines. No direct search URL – would need to automate search form; skipping to avoid fake prices."""
        try:
            print(f"    Scraping PIA for {route['origin']}-{route['destination']}")
            print(f"      (Search not implemented – opening homepage would give wrong numbers; skipping)")
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
            ], route=route)
            
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
            ], route=route)
            
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
            ], route=route)
            
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
            ], route=route)
            
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
        filename = os.path.abspath(filename or self.excel_path)
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
        return (wb, ws, date_col, row, thin_border, avg_fill, flight_header_row)
    
    def _find_existing_rows_for_route(self, ws, flight_header_row: int, route_code: str) -> List[Tuple[int, str, str]]:
        """Find existing data rows for this route. Returns list of (row_idx, col4_value, col5_value)."""
        out = []
        r = flight_header_row + 1
        while r <= ws.max_row:
            code = ws.cell(row=r, column=1).value
            if not code or str(code).strip() != route_code:
                break
            out.append((r, str(ws.cell(row=r, column=4).value or ''), str(ws.cell(row=r, column=5).value or '')))
            r += 1
        return out
    
    def _update_route_date_column(self, ws, route_result: Dict, existing_rows: List[Tuple[int, str, str]], date_col: int) -> None:
        """Fill the date column for existing rows of this route (same order as _write_route_to_sheet: prices then N-averages then Y-averages)."""
        prices = route_result.get('prices', [])
        sorted_prices = sorted(prices, key=lambda x: (x.get('airline', 'Various'), x.get('source', '')))
        airline_avg_groups = {}
        for price_data in prices:
            airline = price_data.get('airline', 'Various')
            if airline != 'Various':
                if airline not in airline_avg_groups:
                    airline_avg_groups[airline] = []
                if price_data.get('price') and price_data.get('price') > 0:
                    airline_avg_groups[airline].append(price_data.get('price'))
        values = []
        for price_data in sorted_prices:
            values.append(price_data.get('price'))
        for airline, price_list in airline_avg_groups.items():
            if len(price_list) > 1:
                values.append(round(sum(price_list) / len(price_list)))
        all_valid = [p.get('price') for p in prices if p.get('price') and p.get('price') > 0]
        if all_valid:
            values.append(round(sum(all_valid) / len(all_valid)))
        for i, (row_idx, _, _) in enumerate(existing_rows):
            if i < len(values) and values[i] is not None:
                ws.cell(row=row_idx, column=date_col).value = values[i]
                ws.cell(row=row_idx, column=date_col).number_format = '0'
    
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
    
    def _expected_rows_count(self, route_result: Dict) -> int:
        """Number of rows we would write for this route (sources + N-averages + Y-averages)."""
        prices = route_result.get('prices', [])
        sorted_prices = sorted(prices, key=lambda x: (x.get('airline', 'Various'), x.get('source', '')))
        n = len(sorted_prices)
        airline_avg_groups = {}
        for p in prices:
            airline = p.get('airline', 'Various')
            if airline != 'Various' and p.get('price'):
                airline_avg_groups.setdefault(airline, []).append(p.get('price'))
        for price_list in airline_avg_groups.values():
            if len(price_list) > 1:
                n += 1
        if any(p.get('price') for p in prices):
            n += 1
        return n

    def append_route_to_excel(self, route_result: Dict, filename: str = None):
        """Append one route's data to the Excel file immediately, or update the date column if rows already exist (so we keep only one block of ~121 rows)."""
        filename = os.path.abspath(filename or self.excel_path)
        wb, ws, date_col, row, thin_border, avg_fill, flight_header_row = self._prepare_excel_for_export(filename)
        route_code = route_result['route']['code']
        existing = self._find_existing_rows_for_route(ws, flight_header_row, route_code)
        expected = self._expected_rows_count(route_result)
        if len(existing) == expected and expected > 0:
            self._update_route_date_column(ws, route_result, existing, date_col)
        else:
            row = self._write_route_to_sheet(ws, route_result, row, date_col, thin_border, avg_fill)
        wb.save(filename)
    
    def export_to_excel(self, results: Dict, filename: str = None):
        """Export all flight prices to Excel. If data was already written per-route, this can re-write or just save column widths."""
        filename = os.path.abspath(filename or self.excel_path)
        print(f"\n{'='*60}")
        print("EXPORTING TO EXCEL")
        print(f"{'='*60}")
        
        if 'error' in results:
            print(f"Error: Cannot export - {results['error']}")
            return False
        
        wb, ws, date_col, row, thin_border, avg_fill, flight_header_row = self._prepare_excel_for_export(filename)
        initial_row = row
        for route_result in results.get('routes', []):
            route_code = route_result['route']['code']
            existing = self._find_existing_rows_for_route(ws, flight_header_row, route_code)
            expected = self._expected_rows_count(route_result)
            if len(existing) == expected and expected > 0:
                self._update_route_date_column(ws, route_result, existing, date_col)
            else:
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


def create_fresh_excel(path: str = None) -> str:
    """Create a fresh Excel file with route block and flight price headers (no data). Use to start over.
    Returns the path saved. Backs up existing file if present."""
    path = os.path.abspath(path or FLIGHT_PRICES_EXCEL)
    if os.path.exists(path):
        backup = path.replace('.xlsx', '_backup_%s.xlsx' % datetime.now().strftime('%Y%m%d_%H%M'))
        try:
            os.rename(path, backup)
            print("Backed up existing file to %s" % backup)
        except OSError:
            os.remove(path)

    scraper = FlightPriceScraper(headless=True)
    routes = scraper._get_default_routes()
    scheduled = scraper._get_scheduled_dates_through_2026()

    wb = Workbook()
    ws = wb.active
    ws.title = FLIGHT_PRICES_SHEET_NAME
    ws.sheet_view.rightToLeft = True
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

    # Route block: row 1 headers, rows 2+ data
    for col, h in enumerate(ROUTE_HEADERS, 1):
        c = ws.cell(row=1, column=col)
        c.value = h
        c.font = Font(bold=True)
        c.fill = header_fill
    for row_idx, r in enumerate(routes, 2):
        ws.cell(row=row_idx, column=1).value = r['code']
        ws.cell(row=row_idx, column=2).value = r['commodity_ar']
        ws.cell(row=row_idx, column=3).value = r['origin']
        ws.cell(row=row_idx, column=4).value = r['origin_code']
        ws.cell(row=row_idx, column=5).value = r['destination']
        ws.cell(row=row_idx, column=6).value = r['destination_code']
        ws.cell(row=row_idx, column=7).value = r['duration_months']

    flight_header_row = len(routes) + 2  # blank row then header
    flight_headers = [
        'Code', 'Commodity', 'الدرجة المقابلة لها في الخطوط (Class equivalent in airlines)',
        'CPI-Flag', 'رمز المصدر (Source Code)', 'وكالات الخطوط (Flight Agencies)'
    ]
    for col_idx, header in enumerate(flight_headers, 1):
        cell = ws.cell(row=flight_header_row, column=col_idx)
        cell.value = header
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
    for col, (header_text, _) in enumerate(scheduled, 7):
        cell = ws.cell(row=flight_header_row, column=col)
        cell.value = header_text
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 60
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 35
    for col in range(7, 7 + len(scheduled)):
        ws.column_dimensions[get_column_letter(col)].width = 15

    wb.save(path)
    print("Created fresh Excel: %s" % path)
    return path


def main():
    """Main function. Use --fresh-excel to create a new empty Excel and exit."""
    import sys
    if '--fresh-excel' in sys.argv or '-f' in sys.argv:
        create_fresh_excel()
        return

    scraper = FlightPriceScraper(headless=False)
    results = scraper.scrape_all()

    json_file = 'flight_prices.json'
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print("\nPrices saved to %s" % json_file)

    scraper.export_to_excel(results)


if __name__ == "__main__":
    main()
