import subprocess
import os
import colorama
import time
from colorama import Fore

def get_windows_dir():
    try:
        return os.environ['SystemRoot'] # OS PATH
    except KeyError:
        print("error: systemroot environment variable not set")
        exit(1)

def check_activation_status(windows_dir): # checks if windows is activated and if not uses exception to print cause of error
    try:
        result = subprocess.run(["cscript", "/nologo", f"{windows_dir}\\System32\\slmgr.vbs", "/xpr"], capture_output=True, text=True)
        if result.returncode == 0:
            return True
        else:
            return False
        
    except Exception as e:
        print(f"error: {e}")
        exit(1)

def read_activation_keys(): # gather keys from activation_keys.txt
    try:
        with open("activation_keys.txt", "r") as f:
            return [line.strip() for line in f.readlines()]
    except FileNotFoundError:
        print("error: activation_keys.txt file not found")
        exit(1)

def activate_windows(windows_dir, activation_key): # activate windows with key
    try:
        activation_result = subprocess.run(["cscript", "/nologo", f"{windows_dir}\\System32\\slmgr.vbs", "/ipk", activation_key], capture_output=True, text=True)
        if activation_result.returncode == 0:
            return True
        else:
            return False
    except Exception as e:
        print(f"error: {e}")
        exit(1)

def remove_key_from_file(activation_keys, activation_key): # if key is tried 5 times and failed all it will be removed
    try:
        with open("activation_keys.txt", "w") as f:
            for key in activation_keys:
                if key!= activation_key:
                    f.write(key + "\n")
    except Exception as e:
        print(f"error: {e}")
        exit(1)

def archive_key(activation_key): # saves removed key if needed
    try:
        with open("archived_keys.txt", "a") as f:
            f.write(activation_key + "\n")
    except Exception as e:
        print(f"error: {e}")
        exit(1)

def main():
    windows_dir = get_windows_dir()
    if check_activation_status(windows_dir):
        print(f"windows is already {Fore.GREEN}activated{Fore.RESET}.")
        time.sleep(2)
        
    else:
        print(f"windows is not {Fore.RED}activated{Fore.RESET}. trying to activate")
        activation_keys = read_activation_keys()
        if not activation_keys:
            print(f"\n{Fore.RED}no activation keys found in activation_keys.txt{Fore.RESET}")
            input(f' press enter to {Fore.LIGHTBLACK_EX}EXIT{Fore.RESET}')
        else:
            print(activation_keys)
            print(" ")
            print(f"{Fore.LIGHTBLACK_EX}==================================={Fore.RESET}\n")
            
            max_tries = 5
            for activation_key in activation_keys:
                for _ in range(max_tries):
                    if activate_windows(windows_dir, activation_key):
                        print(f"windows activated successfully with key {activation_key}!")
                        break
                    else:
                        print(f" [ {Fore.RED}✕{Fore.RESET} ]failed to activate windows with key {activation_key}, trying again.. ({_+1}/{max_tries})\n")
                else:
                    print(" ")
                    print(f"{Fore.LIGHTBLACK_EX}==================================={Fore.RESET}")
                    remove_key_from_file(activation_keys, activation_key)
                    archive_key(activation_key)
                    print(f"\nremoved key {activation_key} from file and archived it after {max_tries} tries.")
                    input(f' press enter to {Fore.LIGHTBLACK_EX}EXIT{Fore.RESET}')

if __name__ == "__main__":
    main()
