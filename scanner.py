import win32com.client
from PIL import Image
from reportlab.pdfgen import canvas
import os
from datetime import datetime
date = datetime.now().strftime("%m-%d_%H_%M")


def scan_to_pdf(output_pdf_path, device_id):
    wia = win32com.client.Dispatch("WIA.CommonDialog")
    device_manager = win32com.client.Dispatch("WIA.DeviceManager")

    # Debugging: List all available devices
    print("Available devices:")
    for info in device_manager.DeviceInfos:
        print(f"ID: {info.DeviceID}, Name: {info.Properties['Name'].Value}")

    # Select the scanner using the device ID
    device = None
    for info in device_manager.DeviceInfos:
        if info.DeviceID == device_id:
            device = info.Connect()
            break

    if device:
        print(f"Selected device: {device.Properties['Name'].Value}")
        device.Properties["Document Handling Select"].Value = 1  # 1 indicates feeder

        images = []
        try:
            while True:
                item = device.Items[1]
                img_file = item.Transfer(FormatID="{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}")

                # Save the scanned image as BMP
                image_path = f'scanned_image_{len(images) + 1}.bmp'
                with open(image_path, 'wb') as f:
                    f.write(img_file.FileData.BinaryData)

                # Open the image and append it to the list
                img = Image.open(image_path)
                images.append(image_path)

                if device.Properties["Document Handling Status"].Value == 0:
                    break
        except Exception as e:
            if "no documents left in the document feeder" not in str(e):
                print(f"Unexpected error during scanning: {e}")

        # Convert BMPs to PDF
        c = canvas.Canvas(output_pdf_path)
        for image_path in images:
            img = Image.open(image_path)
            width, height = img.size
            width, height = width * 0.75, height * 0.75
            c.setPageSize((width, height))
            c.drawImage(image_path, 0, 0, width, height)
            c.showPage()
            img.close()

        c.save()

        # Clean up
        for image_path in images:
            os.remove(image_path)
    else:
        print("No device selected")

# Example usage

