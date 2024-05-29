import win32com.client

def list_scanners():
    wia = win32com.client.Dispatch("WIA.DeviceManager")
    for device in wia.DeviceInfos:
        print(f"ID: {device.DeviceID}, Name: {device.Properties['Name'].Value}")

list_scanners()
