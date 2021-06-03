#!/bin/bash
cd /home/eanderson/RPA
source venv/bin/activate
python3 src/automation/appleseed_tx_plan_due_dates.py
python3 src/automation/appleseed_mha_due_dates.py
deactivate
killall -9 chromedriver
killall -9 chromium-browser
killall -9 chrome
