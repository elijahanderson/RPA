#!/bin/bash
cd /home/eanderson/RPA
source venv/bin/activate
python3 src/automation/abhs_client_services.py
deactivate
killall -9 chromedriver
killall -9 chromium-browser