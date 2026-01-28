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

5. **Extending Routes:**
   - Add new routes in the `_get_routes()` method
   - Each route needs: code, origin, destination, commodity description, etc.
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


class FlightPriceScraper:
    """Scraper for flight prices from multiple airlines and travel websites"""
    
    def __init__(self, headless=False):
        self.headless = headless
        self.driver = None
        self.routes = self._get_routes()
        self.sources = self._get_sources()
        self.results = []
        
    def _get_routes(self) -> List[Dict]:
        """Get list of flight routes to scrape"""
        # Based on the example: Doha-London-Doha, Doha-Cairo-Doha
        # Add more routes as needed
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
        """Calculate departure and return dates (6 months from today)"""
        today = datetime.now()
        departure_date = today + timedelta(days=30)  # 1 month from today
        return_date = departure_date + timedelta(days=months_ahead * 30)  # 6 months later
        
        return departure_date.strftime('%Y-%m-%d'), return_date.strftime('%Y-%m-%d')
    
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
    
    def _extract_price_from_page(self, selectors: List[str], min_price: int = 100, max_price: int = 50000) -> Optional[float]:
        """Extract price from page using multiple selectors"""
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
                # Look for common price patterns
                price_patterns = [
                    r'QAR\s*(\d{1,3}(?:,\d{3})*)',  # QAR followed by number
                    r'(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)',  # Numbers with commas
                    r'(\d{3,6})',  # 3-6 digit numbers
                ]
                for pattern in price_patterns:
                    matches = re.findall(pattern, page_text)
                    if matches:
                        for match in matches[:20]:
                            try:
                                price_val = float(str(match).replace(',', ''))
                                if min_price <= price_val <= max_price:
                                    price = price_val
                                    print(f"      ✓ Extracted price from page source: {price}")
                                    break
                            except:
                                continue
                        if price:
                            break
            except:
                pass
        
        return price
    
    def scrape_qatar_airways(self, route: Dict) -> Optional[Dict]:
        """Scrape prices from Qatar Airways"""
        try:
            print(f"    Scraping Qatar Airways for {route['origin']}-{route['destination']}")
            dep_date, ret_date = self._calculate_dates(route['duration_months'])
            
            # Direct URL with search parameters
            url = (f"https://www.qatarairways.com/app/booking/flight-selection?"
                   f"widget=QR&searchType=F&addTaxToFare=Y&minPurTime=0&selLang=en&"
                   f"tripType=R&fromStation={route['origin_code']}&toStation={route['destination_code']}&"
                   f"departing={dep_date}&returning={ret_date}&bookingClass=E&"
                   f"adults=1&children=0&infants=0&ofw=0&teenager=0&flexibleDate=off&allowRedemption=N")
            
            print(f"      Opening URL: {url[:100]}...")
            self.driver.get(url)
            time.sleep(10)
            
            self._close_dialogs()
            
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
            
            # Direct URL with search parameters
            url = (f"https://www.britishairways.com/nx/b/airselect/en/usa/book/search?"
                   f"trip=round&arrivalDate={ret_date}&departureDate={dep_date}&"
                   f"from={route['origin_code']}&to={route['destination_code']}&"
                   f"travelClass=economy&adults=1&youngAdults=0&children=0&infants=0&bound=outbound")
            
            print(f"      Opening URL: {url[:100]}...")
            self.driver.get(url)
            time.sleep(10)
            
            self._close_dialogs()
            
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
            # KAYAK URL format with sort parameter
            url = f"https://www.kayak.ae/flights/{route['origin_code']}-{route['destination_code']}/{dep_date}/{ret_date}?ucs=bzx8kr&sort=bestflight_a"
            
            print(f"      Opening URL: {url[:100]}...")
            self.driver.get(url)
            time.sleep(10)
            
            self._close_dialogs()
            
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
            
            # eDreams uses hash-based routing
            url = (f"https://www.edreams.qa/travel/#results/"
                   f"type=R;from={route['origin_code']};to={route['destination_code']};"
                   f"dep={dep_date};ret={ret_date};"
                   f"buyPath=FLIGHTS_HOME_SEARCH_FORM;internalSearch=true")
            
            print(f"      Opening URL: {url[:100]}...")
            self.driver.get(url)
            time.sleep(10)
            
            self._close_dialogs()
            
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
            
            url = (f"https://www.cheapoair.com/air/listing?"
                   f"&d1={route['origin_code']}&r1={route['destination_code']}&"
                   f"dt1={dep_date_formatted}&dtype1=A&rtype1=C&"
                   f"d2={route['destination_code']}&r2={route['origin_code']}&"
                   f"dt2={ret_date_formatted}&dtype2=C&rtype2=A&"
                   f"tripType=ROUNDTRIP&cl=ECONOMY&ad=1&se=0&ch=0&infs=0&infl=0")
            
            print(f"      Opening URL: {url[:100]}...")
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
                    "stops": "-1",  # -1 means non-stop only
                    "extraStops": "1",
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
    
    def export_to_excel(self, results: Dict, filename: str = 'flight_prices.xlsx'):
        """Export flight prices to Excel file matching the screenshot format exactly"""
        print(f"\n{'='*60}")
        print("EXPORTING TO EXCEL")
        print(f"{'='*60}")
        
        if 'error' in results:
            print(f"Error: Cannot export - {results['error']}")
            return False
        
        # Create or load workbook
        file_exists = os.path.exists(filename)
        if file_exists:
            print(f"Loading existing file: {filename}")
            wb = load_workbook(filename)
            ws = wb.active
            ws.sheet_view.rightToLeft = True
            max_col = ws.max_column
            if max_col < 7:  # At least 7 columns for headers
                max_col = 7
            print(f"  Existing file has {max_col} columns, {ws.max_row} rows")
        else:
            print(f"Creating new file: {filename}")
            wb = Workbook()
            ws = wb.active
            ws.title = "Flight Prices"
            ws.sheet_view.rightToLeft = True
            
            # Create headers
            headers = [
                'Code',
                'Commodity',
                'الدرجة المقابلة لها في الخطوط (Class equivalent in airlines)',
                'CPI-Flag',
                'رمز المصدر (Source Code)',
                'وكالات الخطوط (Flight Agencies)'
            ]
            
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.value = header
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border
            
            max_col = 6  # 6 header columns
            print(f"  Created headers in {max_col} columns")
        
        # Get current week date
        today = datetime.now()
        date_text = today.strftime('%d-%b')  # Format: 3-Jan, 10-Jan, etc.
        print(f"  Date column: {date_text}")
        
        # Check if this week's column already exists
        date_col = None
        for col in range(7, max_col + 1):
            cell_value = ws.cell(row=1, column=col).value
            if cell_value and date_text in str(cell_value):
                date_col = col
                print(f"  Found existing date column at column {date_col}")
                break
        
        if date_col is None:
            date_col = max_col + 1
            # Add date header with border
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            date_cell = ws.cell(row=1, column=date_col)
            date_cell.value = date_text
            date_cell.font = Font(bold=True)
            date_cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
            date_cell.alignment = Alignment(horizontal='center', vertical='center')
            date_cell.border = thin_border
            print(f"  Created new date column at column {date_col}")
        
        # Ensure header row has yellow fill on ALL columns (1 through date_col)
        header_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        for col in range(1, date_col + 1):
            header_cell = ws.cell(row=1, column=col)
            header_cell.fill = header_fill
            header_cell.font = Font(bold=True)
            header_cell.border = thin_border
            header_cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Find the next available row (if file exists, find last row with data)
        if file_exists:
            # Find the last row with data in column A (Code column)
            row = 2
            while ws.cell(row=row, column=1).value is not None:
                row += 1
            print(f"  Starting from row {row} (after existing data)")
        else:
            row = 2  # Start from row 2 (row 1 is headers)
            print(f"  Starting from row {row} (new file)")
        
        initial_row = row
        total_prices_written = 0
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Light orange/peach color for averages (similar to screenshot)
        avg_fill = PatternFill(start_color='FFE4B5', end_color='FFE4B5', fill_type='solid')
        
        # Process each route
        for route_result in results.get('routes', []):
            route = route_result['route']
            prices = route_result.get('prices', [])
            print(f"  Processing route {route['code']}: {len(prices)} prices found")
            
            route_start_row = row  # Track where this route starts for merging
            
            # Group prices by airline for proper averaging
            airline_price_map = {}
            for price_data in prices:
                airline = price_data.get('airline', 'Various')
                if airline not in airline_price_map:
                    airline_price_map[airline] = []
                airline_price_map[airline].append(price_data)
            
            # Write individual source prices
            # Sort by airline first, then by source
            sorted_prices = sorted(prices, key=lambda x: (
                x.get('airline', 'Various'),
                x.get('source', '')
            ))
            
            for price_data in sorted_prices:
                # Code
                cell = ws.cell(row=row, column=1)
                cell.value = route['code']
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # Commodity (will be merged later)
                cell = ws.cell(row=row, column=2)
                cell.value = route['commodity_ar']
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
                
                # Class
                cell = ws.cell(row=row, column=3)
                cell.value = 'Economy'
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # CPI-Flag
                cell = ws.cell(row=row, column=4)
                cell.value = 'Y'
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # Source Code
                cell = ws.cell(row=row, column=5)
                cell.value = price_data.get('source_code', '')
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # Flight Agencies
                cell = ws.cell(row=row, column=6)
                agency_name = price_data.get('source_ar', price_data.get('source', ''))
                airline_name = price_data.get('airline', '')
                source_name = price_data.get('source', '')
                
                # Format agency name based on source type
                if airline_name and airline_name != 'Various':
                    # For aggregators showing specific airline prices
                    if source_name in ['KAYAK', 'eDreams', 'CheapAir', 'ITA Matrix']:
                        # Format: "Kayak عبر الخطوط العمانية" or "matrix عبر الخطوط العمانية"
                        if source_name == 'KAYAK':
                            cell.value = f"Kayak عبر {airline_name}"
                        elif source_name == 'ITA Matrix':
                            cell.value = f"matrix عبر {airline_name}"
                        else:
                            cell.value = f"{agency_name} عبر {airline_name}"
                    else:
                        # Direct airline booking
                        cell.value = agency_name
                else:
                    # Aggregator without specific airline
                    cell.value = agency_name
                
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='right', vertical='center')
                
                # Price
                price = price_data.get('price')
                cell = ws.cell(row=row, column=date_col)
                if price:
                    cell.value = price
                    cell.number_format = '0'
                    total_prices_written += 1
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                row += 1
            
            # Calculate and add averages for each airline group
            # Group by airline name (not 'Various')
            airline_avg_groups = {}
            for price_data in prices:
                airline = price_data.get('airline', 'Various')
                if airline != 'Various':
                    if airline not in airline_avg_groups:
                        airline_avg_groups[airline] = []
                    price = price_data.get('price')
                    if price and price > 0:
                        airline_avg_groups[airline].append(price)
            
            # Write averages for airlines with multiple prices
            for airline, price_list in airline_avg_groups.items():
                if len(price_list) > 1:  # Only average if multiple sources
                    avg_price = sum(price_list) / len(price_list)
                    
                    # Code
                    cell = ws.cell(row=row, column=1)
                    cell.value = route['code']
                    cell.border = thin_border
                    cell.fill = avg_fill
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    
                    # Commodity (will be merged)
                    cell = ws.cell(row=row, column=2)
                    cell.value = route['commodity_ar']
                    cell.border = thin_border
                    cell.fill = avg_fill
                    cell.alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
                    
                    # Class
                    cell = ws.cell(row=row, column=3)
                    cell.value = 'Economy'
                    cell.border = thin_border
                    cell.fill = avg_fill
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    
                    # CPI-Flag (N-averages or Y-averages)
                    cell = ws.cell(row=row, column=4)
                    cell.value = 'N-averages'
                    cell.border = thin_border
                    cell.fill = avg_fill
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    
                    # Source Code (empty for averages)
                    cell = ws.cell(row=row, column=5)
                    cell.value = ''
                    cell.border = thin_border
                    cell.fill = avg_fill
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    
                    # Flight Agencies (Average) - format: "متوسط المصادر للخطوط [Airline]"
                    cell = ws.cell(row=row, column=6)
                    # Get Arabic airline name if available
                    airline_ar_map = {
                        'Qatar Airways': 'القطرية',
                        'British Airways': 'البريطانية',
                        'Malaysia Airlines': 'الماليزية',
                        'Kuwait Airways': 'الكويتية',
                        'Turkish Airlines': 'التركية',
                        'Pakistan International Airlines': 'الباكستانية',
                        'PIA': 'الباكستانية'
                    }
                    airline_ar = airline_ar_map.get(airline, airline)
                    cell.value = f"متوسط المصادر للخطوط {airline_ar}"
                    cell.border = thin_border
                    cell.fill = avg_fill
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    
                    # Average Price (bold)
                    cell = ws.cell(row=row, column=date_col)
                    cell.value = round(avg_price)
                    cell.number_format = '0'
                    cell.font = Font(bold=True)
                    cell.border = thin_border
                    cell.fill = avg_fill
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    total_prices_written += 1
                    
                    # Fill any columns between 7 and date_col for consistent row fill
                    for c in range(7, date_col):
                        fill_cell = ws.cell(row=row, column=c)
                        fill_cell.fill = avg_fill
                        fill_cell.border = thin_border
                    
                    row += 1
            
            # Add overall average row for ALL prices at the end of each route
            all_valid_prices = [p.get('price') for p in prices if p.get('price') and p.get('price') > 0]
            if len(all_valid_prices) > 0:
                overall_avg = sum(all_valid_prices) / len(all_valid_prices)
                
                # Code
                cell = ws.cell(row=row, column=1)
                cell.value = route['code']
                cell.border = thin_border
                cell.fill = avg_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # Commodity (will be merged)
                cell = ws.cell(row=row, column=2)
                cell.value = route['commodity_ar']
                cell.border = thin_border
                cell.fill = avg_fill
                cell.alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
                
                # Class
                cell = ws.cell(row=row, column=3)
                cell.value = 'Economy'
                cell.border = thin_border
                cell.fill = avg_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # CPI-Flag (Y-averages for overall average)
                cell = ws.cell(row=row, column=4)
                cell.value = 'Y-averages'
                cell.border = thin_border
                cell.fill = avg_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # Source Code (empty for averages)
                cell = ws.cell(row=row, column=5)
                cell.value = ''
                cell.border = thin_border
                cell.fill = avg_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # Flight Agencies (Overall Average)
                cell = ws.cell(row=row, column=6)
                cell.value = "متوسط المصادر"  # Average of all sources
                cell.border = thin_border
                cell.fill = avg_fill
                cell.alignment = Alignment(horizontal='right', vertical='center')
                
                # Overall Average Price (bold)
                cell = ws.cell(row=row, column=date_col)
                cell.value = round(overall_avg)
                cell.number_format = '0'
                cell.font = Font(bold=True)
                cell.border = thin_border
                cell.fill = avg_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
                total_prices_written += 1
                
                # Fill any columns between 7 and date_col for consistent row fill
                for c in range(7, date_col):
                    fill_cell = ws.cell(row=row, column=c)
                    fill_cell.fill = avg_fill
                    fill_cell.border = thin_border
                
                row += 1
            
            # Merge commodity cells for this route (including average rows)
            route_end_row = row - 1
            if route_end_row >= route_start_row:
                try:
                    # Unmerge any existing merged cells in this range first
                    for merged_range in list(ws.merged_cells.ranges):
                        if merged_range.min_col == 2 and merged_range.max_col == 2:
                            if merged_range.min_row <= route_start_row <= merged_range.max_row:
                                ws.unmerge_cells(str(merged_range))
                    
                    # Merge commodity column for this route
                    if route_end_row > route_start_row:
                        ws.merge_cells(f'B{route_start_row}:B{route_end_row}')
                    # Set alignment for merged cell
                    merged_cell = ws.cell(row=route_start_row, column=2)
                    merged_cell.alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
                    print(f"    Merged commodity cells from row {route_start_row} to {route_end_row}")
                except Exception as e:
                    print(f"    Warning: Could not merge commodity cells: {e}")
                    pass  # If merge fails, continue
        
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
        
        try:
            wb.save(filename)
            rows_written = row - initial_row
            print(f"\n✓ Prices exported to {filename} (RTL layout)")
            print(f"  Rows written: {rows_written}")
            print(f"  Prices written: {total_prices_written}")
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
    
    # Export to Excel
    excel_file = 'flight_prices.xlsx'
    scraper.export_to_excel(results, excel_file)


if __name__ == "__main__":
    main()
