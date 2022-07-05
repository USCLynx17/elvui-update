import sys
from os.path import exists
import requests
from zipfile import ZipFile
from winreg import OpenKey, QueryValueEx, HKEY_LOCAL_MACHINE, ConnectRegistry

print()

# Find WoW install directory
location = r"SOFTWARE\WOW6432Node\Blizzard Entertainment\World of Warcraft"
registry = ConnectRegistry(None, HKEY_LOCAL_MACHINE)
wow_addon_path = ""

# Attempt to read from registry to get WoW version
try:
    key = OpenKey(registry, location)
    wow_addon_path = QueryValueEx(key, "InstallPath")[0] + "Interface\\AddOns"
except:
    print("WoW install not found.")
    sys.exit()

# Check ELVUI installed version
elvui_path = wow_addon_path + "\\ElvUI\\ElvUI_Mainline.toc"
elvui_exists = False
elvui_installed_version = ""

if exists(elvui_path):
    elvui_exists = True

if elvui_exists:
    with open(elvui_path, 'r') as file:
        file.readline()
        file.readline()
        file.readline()
        elvui_installed_version = file.readline()[12:17]

# Get ELVUI website content to find latest version
request = requests.get("https://www.tukui.org/download.php?ui=elvui")
content = request.content.decode('UTF-8')
line = content.find("current version")
offset = 47
version_length = 5
start = line + offset
end = start + version_length
elvui_latest_version = content[start:end]

# Download ELVUI
download_file = ""

if elvui_latest_version != elvui_installed_version:
    download_file = "https://www.tukui.org/downloads/elvui-" + elvui_latest_version + ".zip"
    file = requests.get(download_file)
    filename = download_file.split('/')[-1]

    with open(filename, 'wb') as zip:
        zip.write(file.content)

    # Extract zipped file to WoW directory
    with ZipFile(filename, 'r') as zip:
        zip.extractall(wow_addon_path)
else:
    print("Latest version of ELVUI already installed.")

# Debugging
print("WoW Path: " + str(wow_addon_path))
print("ELVUI Installed Version: " + elvui_installed_version)
print("ELVUI Latest Version: " + elvui_latest_version)
print("Download File: " + download_file)