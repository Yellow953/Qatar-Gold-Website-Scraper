# Daily Scheduler Setup

There are several ways to run the gold price scraper daily. Choose the method that works best for your system.

## Method 1: Python Scheduler (Recommended for Testing)

This method uses a Python script that runs continuously and executes the scraper at scheduled times.

### Setup:

1. Make sure you have the schedule library installed:
```bash
pip install -r requirements.txt
```

2. Run the scheduler:
```bash
python scheduler.py
```

3. The scraper will:
   - Run immediately when you start the scheduler
   - Run daily at 9:00 AM (you can change this in `scheduler.py`)
   - Keep running until you stop it (Ctrl+C)

### Customizing the Schedule:

Edit `scheduler.py` and change the time:
```python
# Run at 9:00 AM daily
schedule.every().day.at("09:00").do(run_scraper)

# Or run at multiple times:
schedule.every().day.at("09:00").do(run_scraper)
schedule.every().day.at("15:00").do(run_scraper)  # Also at 3:00 PM

# Or run every 6 hours:
schedule.every(6).hours.do(run_scraper)
```

### Running in Background (Linux/macOS):

To run the scheduler in the background:
```bash
nohup python scheduler.py > scheduler.log 2>&1 &
```

To stop it:
```bash
pkill -f scheduler.py
```

---

## Method 2: Cron Job (Linux/macOS)

Cron is a built-in task scheduler for Linux and macOS.

### Setup:

1. Open your crontab:
```bash
crontab -e
```

2. Add one of these lines (choose based on when you want it to run):

```bash
# Run daily at 9:00 AM
0 9 * * * cd /home/yellow/Documents/coding/python/website-scraper && /usr/bin/python3 gold_scraper.py >> /home/yellow/Documents/coding/python/website-scraper/cron.log 2>&1

# Run daily at 9:00 AM and 3:00 PM
0 9,15 * * * cd /home/yellow/Documents/coding/python/website-scraper && /usr/bin/python3 gold_scraper.py >> /home/yellow/Documents/coding/python/website-scraper/cron.log 2>&1

# Run every 6 hours
0 */6 * * * cd /home/yellow/Documents/coding/python/website-scraper && /usr/bin/python3 gold_scraper.py >> /home/yellow/Documents/coding/python/website-scraper/cron.log 2>&1
```

**Important:** Replace `/home/yellow/Documents/coding/python/website-scraper` with your actual project path, and `/usr/bin/python3` with your Python path (find it with `which python3`).

### Cron Time Format:
```
* * * * *
│ │ │ │ │
│ │ │ │ └─── Day of week (0-7, Sunday = 0 or 7)
│ │ │ └───── Month (1-12)
│ │ └─────── Day of month (1-31)
│ └───────── Hour (0-23)
└─────────── Minute (0-59)
```

### Examples:
- `0 9 * * *` = Every day at 9:00 AM
- `0 9,15 * * *` = Every day at 9:00 AM and 3:00 PM
- `0 */6 * * *` = Every 6 hours
- `30 8 * * 1-5` = Every weekday (Mon-Fri) at 8:30 AM

### Viewing Logs:
```bash
tail -f /home/yellow/Documents/coding/python/website-scraper/cron.log
```

---

## Method 3: Windows Task Scheduler

Windows has a built-in Task Scheduler that can run Python scripts.

### Setup:

1. Open Task Scheduler:
   - Press `Win + R`, type `taskschd.msc`, press Enter

2. Create a new task:
   - Click "Create Basic Task" in the right panel
   - Name: "Gold Price Scraper"
   - Description: "Daily gold price scraper"

3. Set the trigger:
   - Choose "Daily"
   - Set the time (e.g., 9:00 AM)
   - Choose "Recur every: 1 days"

4. Set the action:
   - Choose "Start a program"
   - Program/script: `C:\Python39\python.exe` (or your Python path)
   - Add arguments: `gold_scraper.py`
   - Start in: `C:\path\to\website-scraper` (your project folder)

5. Finish the wizard

### Finding Your Python Path:
```cmd
where python
```

### Alternative: Create a Batch File

Create `run_scraper.bat`:
```batch
@echo off
cd /d "C:\path\to\website-scraper"
python gold_scraper.py
```

Then in Task Scheduler, point to this batch file instead.

---

## Method 4: Systemd Timer (Linux - Modern Alternative)

For modern Linux systems using systemd:

### Create a service file:

1. Create `/etc/systemd/system/gold-scraper.service`:
```ini
[Unit]
Description=Gold Price Scraper
After=network.target

[Service]
Type=oneshot
User=yellow
WorkingDirectory=/home/yellow/Documents/coding/python/website-scraper
ExecStart=/usr/bin/python3 /home/yellow/Documents/coding/python/website-scraper/gold_scraper.py
StandardOutput=append:/home/yellow/Documents/coding/python/website-scraper/scraper.log
StandardError=append:/home/yellow/Documents/coding/python/website-scraper/scraper.log
```

2. Create `/etc/systemd/system/gold-scraper.timer`:
```ini
[Unit]
Description=Run Gold Price Scraper Daily
Requires=gold-scraper.service

[Timer]
OnCalendar=daily
OnCalendar=09:00
Persistent=true

[Install]
WantedBy=timers.target
```

3. Enable and start:
```bash
sudo systemctl daemon-reload
sudo systemctl enable gold-scraper.timer
sudo systemctl start gold-scraper.timer
```

4. Check status:
```bash
sudo systemctl status gold-scraper.timer
```

---

## Recommendations

- **For testing/development**: Use Method 1 (Python scheduler)
- **For Linux/macOS production**: Use Method 2 (Cron) or Method 4 (Systemd)
- **For Windows production**: Use Method 3 (Task Scheduler)

## Troubleshooting

### Check if the script runs manually:
```bash
python gold_scraper.py
```

### Check Python path:
```bash
which python3  # Linux/macOS
where python   # Windows
```

### Check file permissions:
```bash
chmod +x gold_scraper.py
chmod +x scheduler.py
```

### View logs:
- Python scheduler: Check console output or `scheduler.log`
- Cron: Check `cron.log` in project directory
- Systemd: `journalctl -u gold-scraper.service`
