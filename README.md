# atr_tool
Abandoned Treasures Reclaimed Whatnot Seller Tool

This tool was designed to help sellers in the whatnot community compile there  livestream reports into sorted lists by user purchases.

There is definitely a more efficient way to achieve this, but it works, so I haven't bothered. 

# Pyinstaller Windows Command
 - Just copy and paste into terminal

pyinstaller --onefile --windowed --add-data "donate.png;." --add-data "free.png;." --add-data "help.png;." --add-data "howto.png;." --add-data "madeby.png;." --add-data "promo.png;." --add-data "toollogo.png;." --add-data "web_logo.png;." --add-data "whatnot_seller_tool_logo.png;." --add-data "wn.logo;." --add-data "16_3x_XIg_icon.ico;." --icon=16_3x_XIg_icon.ico ATR_Seller_Tool.py

# Pyinstaller Mac Command
 - Coming Soon

# Future Goals
 - Differentiate Local Pickups and Deliveries 
   - Currently Impossible with livestream report document
 - Differentiate completed purchases and cancellations
 - Host on public webserver rather than standalone application

# Extra Info
- Code is encrypted to help download bypass virus scanner on PC's since program is not licensed
- Code originally served one purpose but has been modified to be distributable so there may be a bunch of unnecessary code due to testing various things




