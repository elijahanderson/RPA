cd /home/eanderson/RPA
source venv/bin/activate
python3 src/automation/aacog_csqm.py
deactivate
killall -9 chromedriver
killall -9 chromium-browser
killall -9 chrome
