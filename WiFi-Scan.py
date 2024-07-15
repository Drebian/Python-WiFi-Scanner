# Import Third-Party Modules Necessary
import comtypes
import csv
import io
import os
import pymsgbox
import pywifi
import time

choice = pymsgbox.prompt('Please enter the location that is being scanned: \n N: Northern Atrium \n NE: Northeastern Atrium \n E: Eastern Atrium \n SE: Southeastern Atrium \n S: Southern Atrium \n SW: Southwestern Atrium'
               '\n W: Western Atrium \n NW: Northwestern Atrium \n K: Kitchen \n I: IT Area \n V: Vacations Area \n T: Travelstore Area \n D: Denise Office \n M: Meetings Area \n R: Reception', 'Balboa Travel WiFi Scan')

# Execute Wi-Fi Scan
wifi = pywifi.PyWiFi()
iface = wifi.interfaces()[0]
iface.scan()

timer = 30
while timer > 0:
    pymsgbox.alert('Time Remaining:'+ str(timer), 'Scan in Progress', timeout=1100)
    timer -= 1

results = iface.scan_results()

# Creates the csv file and populates field names in the first row
FilePath = 'C:/Dell/Wifi-Scan.csv'

file_exists = os.path.isfile(FilePath)
       
with io.open(FilePath, 'a', encoding='utf-8', newline='') as file:
    fieldnames = ['BSSID', 'SSID', 'Signal', 'Frequency', 'Authentication Type', 'Encryption', 'Location', ]
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

        freq = network.freq

        auth = network.akm
        if len(auth) == 0:
            auth = "No Authentication"  # Converts blank value to No Authentication

        cipher = network.cipher

        # Converts RSSI to Signal Strength
        strength = int(((network.signal - (-90)) * (100 - 0) / ((-21) - (-90))))

        # Case for defining location
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

        # Case for matching Authentication Type
        match auth:
            case [0]:
                auth = 'None'
            case [1]:
                auth = 'WPA'
            case [2]:
                auth = 'WPA-PSK'
            case [3]:
                auth = 'WPA2'
            case [4]:
                auth = 'WPA2-PSK'
            case [5]:
                auth = 'Unknown'

        #Case for matching Cipher Type
        match cipher:
            case 0:
                cipher = 'None'
            case 1:
                cipher = 'WEP'
            case 2:
                cipher = 'TKIP'
            case 3:
                cipher = 'CCMP'
            case 4:
                cipher = 'Unknown'

        # Writes the scan results into the fields
        row = [bssid, ssid, strength, freq, auth, cipher, location]
        writer.writerow(row)

pymsgbox.alert('Scan Completed.  File saved to ' + str(FilePath), 'Scan Completed', timeout = 10000)
