#!/bin/bash
cd /home/eanderson/RPA
source venv/bin/activate
python3 src/automation/appleseed_tx_plan_due_dates.py
deactivate
