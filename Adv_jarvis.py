import win32com.client as winc
import speech_recognition as sr
import time
import ctypes
import subprocess

def speak(TextToSpeech):
    speaker = winc.Dispatch("SAPI.SpVoice")
    speaker.Speak(TextToSpeech)

def list_removable_devices():
    try:
        c = winc.Dispatch("WbemScripting.SWbemLocator")
        wmi = c.ConnectServer(".", r"root\cimv2")
        
        removable_disks = wmi.ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE DriveType = 2")
        
        devices = []
        if len(removable_disks) > 0:
            speak(f"I have detected {len(removable_disks)} removable {'device' if len(removable_disks) == 1 else 'devices'}.")
            print(f"Detected {len(removable_disks)} removable {'device' if len(removable_disks) == 1 else 'devices'}:")
            
            for index, disk in enumerate(removable_disks, start=1):
                device_info = f"{index}. {disk.VolumeName} ({disk.DeviceID})"
                speak(f"Device {index} found: {disk.VolumeName}")
                print(device_info)
                print(f"   Free Space: {float(disk.FreeSpace) / (1024**3):.2f} GB")
                print(f"   Total Size: {float(disk.Size) / (1024**3):.2f} GB")
                devices.append((index, disk.DeviceID))
        else:
            speak("No removable devices are currently detected.")
            print("No removable devices detected.")
        
        return devices
    
    except Exception as e:
        speak("An error occurred while checking removable devices.")
        print(f"Error: {e}")
        return []

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

def Listen():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=0.5)
        print("Listening...")
        try:
            audio = r.listen(source, timeout=5, phrase_time_limit=5)
        except sr.WaitTimeoutError:
            print("Listening timed out. Please try again.")
            return None
    try:
        query = r.recognize_google(audio, language="en-in")
        print(f"User said: {query}")
        return query.lower()  # Convert the query to lowercase for easier matching
    except sr.UnknownValueError:
        print("Sorry, I did not understand that.")
        return None
    except sr.RequestError:
        print("Could not request results from Google Speech Recognition service")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None
def is_antivirus_on():
    try:
        result = subprocess.run(
            ["powershell", "-Command", "Get-MpComputerStatus | Select-Object -ExpandProperty AMServiceEnabled"],
            capture_output=True, text=True
        )
        
        if result.returncode == 0:
            status = result.stdout.strip()
            if status == "True":
                print("Antivirus is ON.")
                return True
            else:
                print("Antivirus is OFF.")
                return False
        else:
            print("Could not determine antivirus status.")
            return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

# Turn Antivirus On
def turn_antivirus_on():
    try:
        subprocess.run(
            ["powershell", "-Command", "Set-MpPreference -DisableRealtimeMonitoring $false"],
            check=True
        )
        print("Antivirus has been turned ON.")
        speak("Antivirus has been turned ON.")
    except subprocess.CalledProcessError as e:
        print(f"Failed to turn on antivirus: {e}")

# Turn Antivirus Off
def turn_antivirus_off():
    try:
        subprocess.run(
            ["powershell", "-Command", "Set-MpPreference -DisableRealtimeMonitoring $true"],
            check=True
        )
        print("Antivirus has been turned OFF.")
        speak("Antivirus has been turned OFF.")
    except subprocess.CalledProcessError as e:
        print(f"Failed to turn off antivirus: {e}")

def main():
    # Check current antivirus status
    status = is_antivirus_on()
    
    if status is not None:
        speak("Would you like to turn the antivirus on, off, or leave it as it is?")
        user_input = Listen()
        
        if user_input is not None:
            if "on" in user_input and not status:
                turn_antivirus_on()
            elif "off" in user_input and status:
                turn_antivirus_off()
            elif "leave" in user_input:
                print("No changes made to antivirus status.")
                speak("No changes made to antivirus status.")
            else:
                print("Antivirus is already in the requested state, or invalid command.")
                speak("Antivirus is already in the requested state, or I did not understand the command.")
        else:
            print("Could not understand your response.")
            speak("I could not understand your response.")
    else:
        print("Could not retrieve antivirus status.")
        speak("Could not retrieve antivirus status.")
# Main program execution
speak("Hello, I am JARVIS A.I.")
speak("I'm ready to assist you.")

devices = []

while True:
    speak("Please give your order, sir.")
    text = Listen()
    speak("recognizing")
    print("recognizing")
    if text is None:
        continue
    
    if "check" in text or "list" in text:
        speak("Checking for removable devices.")
        devices = list_removable_devices()
    elif "eject device" in text:
        if not devices:
            speak("Please check for devices first.")
            continue
        
        speak("Which device number would you like to eject?")
        device_number = int(input("Enter the device number:"))
        if device_number:
            device_number = int(device_number)
            matching_devices = [dev for dev in devices if dev[0] == device_number]
            if matching_devices:
                eject_device(matching_devices[0][1])
            else:
                speak(f"I'm sorry, device number {device_number} is not in the list.")
        else:
            speak("I'm sorry, I couldn't understand the device number. Please try again.")
    elif "antivirus" in text:
        speak("checking for the status")
        main()
    elif "exit" in text or "close" in text:
        speak("Closing now, sir.")
        break
    else:
        speak("I'm sorry, I didn't understand that command. Could you please repeat?")

    time.sleep(1)  # Add a small delay to prevent rapid looping
