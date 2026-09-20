"""
Sideloads the Excel add-in using the WEF\Developer registry approach.
This is the same method used by Microsoft's official office-addin-debugging tool.
No file share, no admin, no catalog needed.
"""
import winreg, sys, os, xml.etree.ElementTree as ET

MANIFEST_PATH = r"C:\Users\Asus\Desktop\COMMANDCODE PROJECTS\excel-ai-addin\manifest.xml"

# Read the add-in GUID from the manifest
tree = ET.parse(MANIFEST_PATH)
ns = {"o": "http://schemas.microsoft.com/office/appforoffice/1.1"}
addin_id = tree.find("o:Id", ns).text.strip()
print(f"Add-in ID from manifest: {addin_id}")

success = False
for ver in ["16.0", "15.0"]:
    try:
        DEV_KEY = f"Software\\Microsoft\\Office\\{ver}\\WEF\\Developer"

        # Create or open the Developer key
        dev = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, DEV_KEY, 0, winreg.KEY_ALL_ACCESS)

        # Use add-in GUID as the subkey name (matches office-addin-debugging behavior)
        subkey_name = addin_id.upper()

        # Delete old entry if exists
        try:
            winreg.DeleteKey(dev, subkey_name)
            print(f"[{ver}] Removed old developer entry.")
        except FileNotFoundError:
            pass

        entry = winreg.CreateKeyEx(dev, subkey_name, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(entry, "ManifestPath", 0, winreg.REG_SZ,    MANIFEST_PATH)
        winreg.SetValueEx(entry, "UseDirectUrl",  0, winreg.REG_DWORD, 1)
        winreg.CloseKey(entry)
        winreg.CloseKey(dev)

        print(f"[{ver}] Developer sideload registered: {MANIFEST_PATH}")
        success = True
    except Exception as e:
        print(f"[{ver}] Failed: {e}")

if success:
    print()
    print("Done! Steps:")
    print("  1. Fully close and reopen Excel")
    print("  2. Insert -> My Add-ins -> MY ADD-INS tab (or Developer tab -> Add-ins)")
    print("  3. You should see 'AI Assistant' listed there")
    print()
    print("  If not visible: Insert -> My Add-ins -> Refresh button")
else:
    print("Failed - see errors above.")
    sys.exit(1)
