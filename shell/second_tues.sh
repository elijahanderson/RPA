#!/bin/bash
cd /home/eanderson/RPA
source venv/bin/activate
python3 src/automation/oaks_crisis.py
deactivate
killall -9 chromedriver
killall -9 chromium-browser
