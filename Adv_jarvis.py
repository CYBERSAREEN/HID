import win32com.client as winc
import ctypes
import subprocess

def speak(TextToSpeech):
    speaker = winc.Dispatch("SAPI.SpVoice")
    speaker.Speak(TextToSpeech)

def list_devices():
    try:
        c = winc.Dispatch("WbemScripting.SWbemLocator")
        wmi = c.ConnectServer(".", r"root\cimv2")
        
        # Query for removable disks
        removable_disks = wmi.ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE DriveType = 2")
        
        # Query for HID devices
        hid_devices = wmi.ExecQuery("SELECT * FROM Win32_PnPEntity WHERE PNPClass = 'HIDClass' OR PNPClass = 'USB'")
        
        device_count = len(removable_disks) + len(hid_devices)
        
        if device_count > 0:
            speak(f"I have detected {device_count} {'device' if device_count == 1 else 'devices'}.")
            print(f"Detected {device_count} {'device' if device_count == 1 else 'devices'}:")
            
            # List removable disks
            if len(removable_disks) > 0:
                print("\nRemovable Storage Devices:")
                for disk in removable_disks:
                    speak(f"Removable storage device found: {disk.VolumeName} ({disk.DeviceID})")
                    print(f"- {disk.VolumeName} ({disk.DeviceID})")
                    print(f"  Free Space: {disk.FreeSpace / (1024**3):.2f} GB")
                    print(f"  Total Size: {disk.Size / (1024**3):.2f} GB")
            
            # List HID devices
            if len(hid_devices) > 0:
                print("\nHID Devices:")
                for device in hid_devices:
                    speak(f"HID device found: {device.Name}")
                    print(f"- {device.Name}")
                    print(f"  Device ID: {device.DeviceID}")
        else:
            speak("No devices are currently detected.")
            print("No devices detected.")
    
    except Exception as e:
        speak("An error occurred while checking devices.")
        print(f"Error: {e}")

def eject_device(drive_letter):
    try:
        drive_letter = drive_letter.upper()[:1]
        
        winapi = ctypes.windll.kernel32
        
        drive_name = fr"\\.\{drive_letter}:"
        
        handle = winapi.CreateFileW(drive_name, 0x80000000 | 0x40000000, 0x1 | 0x2, None, 0x3, 0, None)
        if handle == -1:
            raise ctypes.WinError()
        
        result = winapi.DeviceIoControl(handle, 0x2D4808, None, 0, None, 0, None, None)
        if result == 0:
            raise ctypes.WinError()
        
        winapi.CloseHandle(handle)
        
        speak(f"Device {drive_letter} has been safely ejected.")
        print(f"Device {drive_letter} has been safely ejected.")
    except Exception as e:
        speak(f"An error occurred while trying to eject device {drive_letter}.")
        print(f"Error ejecting device {drive_letter}: {e}")

def detach_hid_device(device_id):
    try:
        # Use pnputil to disable the device
        result = subprocess.run(['pnputil', '/disable-device', device_id], capture_output=True, text=True, check=True)
        
        if "Device disabled successfully." in result.stdout:
            speak(f"HID device with ID {device_id} has been successfully detached.")
            print(f"HID device with ID {device_id} has been successfully detached.")
        else:
            speak(f"Failed to detach HID device with ID {device_id}.")
            print(f"Failed to detach HID device with ID {device_id}.")
            print(f"Output: {result.stdout}")
    except subprocess.CalledProcessError as e:
        speak(f"An error occurred while trying to detach HID device {device_id}.")
        print(f"Error detaching HID device {device_id}: {e}")
        print(f"Error output: {e.stderr}")

# Main program execution
speak("Hello, I am JARVIS A.I.")
print("Hello, I am JARVIS A.I.")
speak("I'm ready to assist you.")
print("I'm ready to assist you.")

while True:
    print("\nAvailable commands:")
    print("1. list - List all devices (removable storage and HID)")
    print("2. eject <drive_letter> - Eject a specific removable storage device (e.g., eject E)")
    print("3. detach <device_id> - Detach a specific HID device (e.g., detach HID\\VID_16C0&PID_27DB\\6&13C452C7&0&2)")
    print("4. exit - Close the program")
    
    command = input("Enter your command: ").strip().lower()
    
    if command == "list":
        speak("Checking for devices.")
        list_devices()
    elif command.startswith("eject"):
        parts = command.split()
        if len(parts) == 2 and len(parts[1]) == 1:
            eject_device(parts[1])
        else:
            speak("Invalid eject command. Please use the format 'eject <drive_letter>'.")
            print("Invalid eject command. Please use the format 'eject <drive_letter>'.")
    elif command.startswith("detach"):
        parts = command.split(maxsplit=1)
        if len(parts) == 2:
            detach_hid_device(parts[1])
        else:
            speak("Invalid detach command. Please use the format 'detach <device_id>'.")
            print("Invalid detach command. Please use the format 'detach <device_id>'.")
    elif command == "exit":
        speak("Closing now, sir.")
        print("Closing now, sir.")
        break
    else:
        speak("I'm sorry, I didn't understand that command. Please try again.")
        print("I'm sorry, I didn't understand that command. Please try again.")

speak("Thank you for using JARVIS A.I! Goodbye sir!")