# Whatnot Seller Tool
Made by Abandoned Treasures Reclaimed

This tool was designed to help sellers in the whatnot community compile there livestream reports into aggregated lists by unique user purchases.

~~There is definitely a more efficient way to achieve this, but it works, so I haven't bothered.~~
***I bothered to do so...***

# What does it do?
This program takes the livestream report(.csv) that is generated at the end of a Whatnot show as input and will output a file name "PRINT ME.docx"
which contains a list of all unique buyers and their subsequent live purchases, buy it now purchases, and giveaways won.

# Pyinstaller Windows Command
 - Just copy and paste into terminal to create a single executable application

pyinstaller --onefile --windowed --add-data "gui.png;." --add-data "16_3x_XIg_icon.ico;." --icon=16_3x_XIg_icon.ico ATR_Seller_Tool.py

# Pyinstaller Mac Command
 - Coming Soon

# Future Goals
 - Differentiate Local Pickups and Deliveries 
   - Currently Impossible with livestream report document
 - ~~Differentiate completed purchases and cancellations~~
 - Host on public webserver rather than standalone application
 - ~~Track coupon data~~






