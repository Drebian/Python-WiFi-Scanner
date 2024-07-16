# Import Third-Party Modules Necessary
import comtypes
import csv
import io
import os
import pymsgbox
import pywifi
from pywifi import _wifiutil_win
import re
import time

# Prompts for location from user
choice = pymsgbox.prompt('Please enter the location that is being scanned: \n N: Northern Atrium \n NE: Northeastern Atrium \n E: Eastern Atrium \n SE: Southeastern Atrium \n S: Southern Atrium \n SW: Southwestern Atrium'
               '\n W: Western Atrium \n NW: Northwestern Atrium \n K: Kitchen \n I: IT Area \n V: Vacations Area \n T: Travelstore Area \n D: Denise Office \n M: Meetings Area \n R: Reception', 'Balboa Travel WiFi Scan')

# Populates location field based on user input
match choice:
    case 'N':
        location = "Northern Atrium"
    case 'NE':
        location = "Northeastern Atrium"
    case 'E':
        location = "Eastern Atrium"
    case 'SE':
        location = "Southeastern Atrium"
    case 'S':
        location = "Southern Atrium"
    case 'SW':
        location = "Southwestern Atrium"
    case 'W':
        location = "Western Atrium"
    case 'NW':
        location = "Northwestern Atrium"
    case 'K':
        location = "Kitchen"
    case 'I':
        location = "IT Area"
    case 'V':
        location = "Vacations"
    case 'T':
        location = "Travelstore Area"
    case 'D':
        location = "Denise's Office"
    case 'M':
        location = "Meetings Area"
    case 'R':
        location = "Reception"
                
# Execute Wi-Fi Scan
wifi = pywifi.PyWiFi()
iface = wifi.interfaces()[0]
iface.scan()

timer = 30
while timer > 0:
    pymsgbox.alert('Time Remaining: ' + str(timer) + ' seconds', 'Scan in Progress', timeout=1000)
    timer -= 1

results = iface.scan_results()

# Creates the csv file and populates field names in the first row
FilePath = 'C:/Dell/Wifi-Scan.csv'

file_exists = os.path.isfile(FilePath)
       
with io.open(FilePath, 'a', encoding='utf-8', newline='') as file:
    fieldnames = ['BSSID', 'SSID', 'Signal', 'Authentication Type', 'Encryption', 'Location']
    writer = csv.writer(file)

    # If File already exists, do not write headers.
    if not file_exists:
        writer.writerow(fieldnames)

    #Creates the link between fieldnames and variables
    for network in results:
        bssid = network.bssid[:-1]

        ssid = network.ssid
        if len(ssid) == 0:
            ssid = "No SSID"  # Converts blank value to No SSID

        auth = _wifiutil_win.WifiUtil._get_auth_alg
        cipher = network.akm

        # Converts RSSI to Signal Strength
        strength = int(((network.signal - (-90)) * (100 - 0) / ((-21) - (-90))))

        # Case for matching Authentication Type
        match auth:
            case [0]:
                auth = 'None'
            case [1]:
                auth = 'Open'
            case [2]:
                auth = 'Shared Key'
            case [3]:
                auth = 'WPA'
            case [4]:
                auth = 'WPA-PSK'
            case [5]:
                auth = 'WPA None'
            case [6]:
                auth = 'WPA2'
            case [7]:
                auth = 'WPA2-PSK'
            case [8]:
                auth = 'WPA3'
            case [9]:
                auth = 'WPA3 Enterprise'
            case [10]:
                auth = 'WPA3 SAE'
            case [11]:
                auth = 'WPA3 Enterprise'

        #Case for matching Cipher Type
        match cipher:
            case [0]:
                cipher = 'None'
            case [1]:
                cipher = 'WEP'
            case [2]:
                cipher = 'TKIP'
            case [4]:
                cipher = 'CCMP'
            case [5]:
                cipher = 'WEP'
            case [100]:
                cipher = 'WPA'
            case [101]:
                cipher = 'WEP'

        # Writes the scan results into the fields
        row = [bssid, ssid, strength, auth, cipher, location]
        writer.writerow(row)

pymsgbox.alert('Scan Completed.  File saved to ' + str(FilePath), 'Scan Completed')
