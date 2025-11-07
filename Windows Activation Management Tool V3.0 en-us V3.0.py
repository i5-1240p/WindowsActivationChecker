# -*- coding: utf-8 -*-
import subprocess
import re
import platform
import ctypes
import sys
import os
import time
import locale
import winreg

class WindowsActivationManager:
    def __init__(self):
        self.admin = self.is_admin()
        self.running = True
        self.setup_encoding()
        self.show_disclaimer()
        
    def setup_encoding(self):
        """Set encoding to solve garbled text issues"""
        try:
            os.system('chcp 65001 > nul')
            if hasattr(sys.stdout, 'reconfigure'):
                sys.stdout.reconfigure(encoding='utf-8')
        except:
            pass
        self.encoding = locale.getpreferredencoding() or 'gbk'
        
    def show_disclaimer(self):
        """Display disclaimer"""
        self.clear_screen()
        print("=" * 70)
        print("                   DISCLAIMER")
        print("=" * 70)
        print()
        print("Important Notice:")
        print("1. This program is for educational and research purposes only")
        print("2. Please ensure you have a legitimate Windows license")
        print("3. KMS activation is only applicable to environments with legitimate volume licenses")
        print("4. Illegal activation of Windows may violate laws and regulations")
        print("5. Users are solely responsible for their actions")
        print("6. Developers assume no legal responsibility for misuse of this program")
        print()
        print("Continuing to use this program indicates you have read and agree to the above terms")
        print()
        
        try:
            input("Press Enter to accept disclaimer and continue...")
        except:
            pass
    
    def is_admin(self):
        """Check if running with administrator privileges"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False
    
    def clear_screen(self):
        """Clear screen"""
        os.system('cls' if os.name == 'nt' else 'clear')
    
    def print_header(self):
        """Print header"""
        self.clear_screen()
        print("=" * 60)
        print("       Windows Activation Management Tool")
        print("=" * 60)
        print()
        
        if self.admin:
            print("[√] Running with administrator privileges")
        else:
            print("[!] Not running with administrator privileges, some features may be limited")
        print()
    
    def get_system_tool_path(self, tool_name):
        """Get full path of system tools"""
        windir = os.environ.get('WINDIR', 'C:\\Windows')
        possible_paths = [
            os.path.join(windir, 'System32', tool_name),
            os.path.join(windir, 'SysWOW64', tool_name),
            os.path.join(windir, 'system32', tool_name),
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                return path
        return tool_name
    
    def run_command(self, cmd, show_output=True):
        """Execute command and return output"""
        try:
            result = subprocess.run(
                cmd, 
                capture_output=True, 
                text=True, 
                shell=True, 
                encoding=self.encoding,
                errors='replace'
            )
            output = result.stdout
            if show_output and output.strip():
                for line in output.split('\n'):
                    if line.strip():
                        print(f"  {line.strip()}")
            return output
        except Exception as e:
            error_msg = f"Command execution error: {str(e)}"
            if show_output:
                print(f"  [Error] {error_msg}")
            return error_msg
    
    def run_slmgr_command(self, args, show_output=True):
        """Run slmgr command"""
        slmgr_path = self.get_system_tool_path('slmgr.vbs')
        cscript_path = self.get_system_tool_path('cscript.exe')
        
        cmd = f'"{cscript_path}" //nologo "{slmgr_path}" {args}'
        return self.run_command(cmd, show_output)
    
    def run_wmic_command(self, args, show_output=True):
        """Run WMIC command"""
        wmic_path = self.get_system_tool_path('wmic.exe')
        cmd = f'"{wmic_path}" {args}'
        return self.run_command(cmd, show_output)
    
    def run_powershell_command(self, args, show_output=True):
        """Run PowerShell command"""
        powershell_path = self.get_system_tool_path('powershell.exe')
        cmd = f'"{powershell_path}" {args}'
        return self.run_command(cmd, show_output)
    
    def check_activation_status(self):
        """Check activation status"""
        print("[Activation Status Check]")
        
        # Check using slmgr command
        print("1. Check using slmgr command:")
        self.run_slmgr_command("/dli")
        
        print("\n2. License details:")
        self.run_slmgr_command("/dlv")
        
        print("\n3. Activation expiration:")
        self.run_slmgr_command("/xpr")
        
        # Get basic info using WMIC
        print("\n4. System information:")
        try:
            output = self.run_wmic_command('os get caption,version /format:value', False)
            if output:
                for line in output.split('\n'):
                    if '=' in line:
                        key, value = line.split('=', 1)
                        print(f"  {key.strip()}: {value.strip()}")
        except:
            pass
        
        print()
    
    def get_oem_info(self):
        """Get OEM information - Enhanced version"""
        print("[OEM Information Query]")
        
        found_key = False
        
        # Method 1: Get OA3xOriginalProductKey using PowerShell
        print("1. Query OEM key using PowerShell:")
        try:
            ps_cmd = '(Get-WmiObject -Query "SELECT * FROM SoftwareLicensingService").OA3xOriginalProductKey'
            output = self.run_powershell_command(f'-Command "{ps_cmd}"', False)
            if output and len(output.strip()) > 10:
                key = output.strip()
                print(f"  OEM Product Key: {key}")
                found_key = True
            else:
                print("  OEM product key not found (Method 1)")
        except Exception as e:
            print(f"  PowerShell query failed: {e}")
        
        # Method 2: Query OEM key using WMIC
        if not found_key:
            print("\n2. Query OEM key using WMIC:")
            try:
                output = self.run_wmic_command('path softwarelicensingservice get OA3xOriginalProductKey /format:value', False)
                if output and 'OA3xOriginalProductKey' in output:
                    for line in output.split('\n'):
                        if 'OA3xOriginalProductKey=' in line:
                            key = line.split('=', 1)[1].strip()
                            if len(key) > 10:
                                print(f"  OEM Product Key: {key}")
                                found_key = True
                                break
                if not found_key:
                    print("  OEM product key not found (Method 2)")
            except Exception as e:
                print(f"  WMIC query failed: {e}")
        
        # Method 3: Read from registry
        if not found_key:
            print("\n3. Query OEM information from registry:")
            try:
                # Try to read DigitalProductId
                key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion"
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                        try:
                            product_id, _ = winreg.QueryValueEx(key, "ProductId")
                            if product_id and len(product_id) > 10:
                                print(f"  Product ID: {product_id}")
                        except FileNotFoundError:
                            pass
                        
                        try:
                            digital_product_id, _ = winreg.QueryValueEx(key, "DigitalProductId")
                            if digital_product_id:
                                print("  Found DigitalProductId (needs decoding)")
                        except FileNotFoundError:
                            pass
                except Exception as e:
                    print(f"  Registry read failed: {e}")
            except Exception as e:
                print(f"  Registry query failed: {e}")
        
        # Method 4: Use system information command
        print("\n4. System manufacturer information:")
        try:
            # Get computer manufacturer and model
            output = self.run_wmic_command('computersystem get manufacturer,model /format:value', False)
            if output:
                for line in output.split('\n'):
                    if '=' in line:
                        key, value = line.split('=', 1)
                        print(f"  {key.strip()}: {value.strip()}")
        except Exception as e:
            print(f"  Failed to get system information: {e}")
        
        # Method 5: Get BIOS information
        print("\n5. BIOS Information:")
        try:
            output = self.run_wmic_command('bios get manufacturer,serialnumber,version /format:value', False)
            if output:
                for line in output.split('\n'):
                    if '=' in line:
                        key, value = line.split('=', 1)
                        print(f"  {key.strip()}: {value.strip()}")
        except Exception as e:
            print(f"  Failed to get BIOS information: {e}")
        
        if not found_key:
            print("\n[Note] OEM product key not found, possible reasons:")
            print("  - This device is not an OEM device")
            print("  - System has been reinstalled, OEM information lost")
            print("  - Insufficient current user privileges")
        
        print()

    def show_installed_product_keys(self):
        """Display installed product keys"""
        print("[Installed Product Keys]")
        print()
        
        # Method 1: Display key info using slmgr command
        print("1. Query using slmgr command:")
        output = self.run_slmgr_command("/dli", False)
        
        # Try to extract key information from output
        lines = output.split('\n')
        key_found = False
        
        for line in lines:
            line_lower = line.lower()
            if 'product key' in line_lower or 'partial product key' in line_lower:
                print(f"  {line.strip()}")
                key_found = True
        
        # Method 2: Query detailed info using PowerShell
        print("\n2. Query using PowerShell:")
        try:
            ps_cmd = 'Get-WmiObject -Query "SELECT * FROM SoftwareLicensingService" | Select-Object OA3xOriginalProductKey, Version, MachineName'
            output = self.run_powershell_command(f'-Command "{ps_cmd}"', False)
            if output and 'OA3xOriginalProductKey' in output:
                for line in output.split('\n'):
                    if line.strip() and not line.strip().startswith('OA3x'):
                        print(f"  {line.strip()}")
                        key_found = True
        except Exception as e:
            print(f"  PowerShell query failed: {e}")
        
        # Method 3: Query Windows installed keys
        print("\n3. Query Windows installed keys:")
        try:
            ps_cmd = '(Get-WmiObject -Query "SELECT * FROM SoftwareLicensingProduct WHERE PartialProductKey <> null").PartialProductKey'
            output = self.run_powershell_command(f'-Command "{ps_cmd}"', False)
            if output and output.strip():
                keys = [k.strip() for k in output.split('\n') if k.strip()]
                for key in keys:
                    if len(key) > 5:
                        print(f"  Partial Product Key: {key}*****")
                        key_found = True
        except Exception as e:
            print(f"  Query installed keys failed: {e}")
        
        # Method 4: Read from registry
        print("\n4. Query from registry:")
        try:
            key_paths = [
                r"SOFTWARE\Microsoft\Windows NT\CurrentVersion",
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\OOBE"
            ]
            
            for key_path in key_paths:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                        i = 0
                        while True:
                            try:
                                name, value, type_ = winreg.EnumValue(key, i)
                                if any(kw in name.lower() for kw in ['key', 'productkey', 'digitalproductid']):
                                    if value and len(str(value)) > 10:
                                        # Partially hide actual keys
                                        if 'productkey' in name.lower() and len(str(value)) > 15:
                                            masked_value = str(value)[:20] + "*****"
                                            print(f"  {name}: {masked_value}")
                                        else:
                                            print(f"  {name}: {value}")
                                        key_found = True
                                i += 1
                            except WindowsError:
                                break
                except FileNotFoundError:
                    pass
                except Exception as e:
                    print(f"  Failed to read registry path {key_path}: {e}")
        except Exception as e:
            print(f"  Registry query failed: {e}")
        
        if not key_found:
            print("  No installed product key information found")
        
        print("\n[Note]")
        print("  - For security reasons, complete keys are usually not displayed directly")
        print("  - If partial keys (*****) are shown, it means keys exist in system but are hidden")
        print("  - OEM devices typically have BIOS-embedded keys")
        print()

    def validate_product_key(self, key):
        """Validate product key format"""
        pattern = r'^[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}-[A-Z0-9]{5}$'
        return re.match(pattern, key.upper()) is not None

    def install_product_key(self):
        """Install product key"""
        print("[Install Product Key]")
        
        if not self.admin:
            print("[Error] This operation requires administrator privileges, please run the program as administrator")
            print()
            return False
        
        print("Please enter a valid Windows product key (format: XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)")
        print("Or press Enter directly to cancel operation")
        print()
        
        while True:
            try:
                product_key = input("Product Key: ").strip().upper()
                
                if not product_key:
                    print("Installation cancelled")
                    return False
                    
                if self.validate_product_key(product_key):
                    break
                else:
                    print("Invalid key format, please re-enter")
                    print()
            except UnicodeDecodeError:
                print("Input encoding error, please re-enter")
            except Exception as e:
                print(f"Input error: {str(e)}")
        
        print()
        print(f"Key to be installed: {product_key[:24]}*****")
        try:
            confirm = input("Confirm installation? (y/N): ").strip().lower()
        except:
            confirm = 'n'
        
        if confirm != 'y':
            print("Installation cancelled")
            return False
        
        print()
        print("Installing product key...")
        result = self.run_slmgr_command(f"/ipk {product_key}", False)
        
        if "success" in result or "successfully" in result.lower():
            print("[Success] Product key installed successfully")
            return True
        else:
            print(f"[Error] Product key installation failed: {result}")
            return False

    def activate_windows(self):
        """Activate Windows"""
        print("[Activate Windows]")
        
        if not self.admin:
            print("[Error] This operation requires administrator privileges, please run the program as administrator")
            print()
            return False
        
        try:
            confirm = input("Confirm to attempt Windows activation? (y/N): ").strip().lower()
        except:
            confirm = 'n'
            
        if confirm != 'y':
            print("Activation cancelled")
            return False
        
        print()
        print("Attempting to activate Windows...")
        result = self.run_slmgr_command("/ato", False)
        
        if "success" in result or "successfully" in result.lower() or "activated" in result:
            print("[Success] Windows activated successfully")
            return True
        else:
            print(f"[Warning] Activation result: {result}")
            return False
    
    def kms_activation(self):
        """KMS activate Windows"""
        print("[KMS Activation]")
        print()
        
        # Display KMS activation disclaimer
        print("Important Reminder:")
        print("1. KMS activation is only applicable to environments with legitimate volume licenses")
        print("2. Please ensure you have legitimate rights to use KMS activation")
        print("3. Illegal use of KMS activation may violate laws and regulations")
        print("4. Using this feature means you understand and assume corresponding responsibility")
        print()
        
        try:
            confirm = input("I have read and understood the above reminder, confirm to continue? (y/N): ").strip().lower()
        except:
            confirm = 'n'
            
        if confirm != 'y':
            print("KMS activation cancelled")
            return False
        
        if not self.admin:
            print("[Error] This operation requires administrator privileges, please run the program as administrator")
            print()
            return False
        
        print()
        print("Please select KMS activation option:")
        print("1. Custom KMS activation (enter KMS server address)")
        print("2. Cancel")
        print()
        
        try:
            choice = input("Please select (1-2): ").strip()
        except:
            choice = '2'
        
        if choice == '1':
            return self.custom_kms_activation()
        else:
            print("KMS activation cancelled")
            return False
    
    def custom_kms_activation(self):
        """Custom KMS activation"""
        print()
        print("Please enter KMS server address (example: kms.example.com)")
        print("Or press Enter directly to cancel")
        print()
        
        try:
            kms_server = input("KMS Server: ").strip()
        except:
            kms_server = ""
        
        if not kms_server:
            print("KMS activation cancelled")
            return False
        
        # Validate server format
        if not re.match(r'^[a-zA-Z0-9.-]+$', kms_server):
            print("[Error] Invalid server address format")
            return False
        
        print()
        print(f"Using KMS server: {kms_server}")
        print("Setting KMS server...")
        
        # Set KMS server
        result = self.run_slmgr_command(f"/skms {kms_server}", False)
        
        print("Attempting activation...")
        activate_result = self.run_slmgr_command("/ato", False)
        
        if "success" in activate_result or "successfully" in activate_result.lower() or "activated" in activate_result:
            print("[Success] KMS activation successful")
            return True
        else:
            print(f"[Error] KMS activation failed: {activate_result}")
            return False
    
    def reset_windows_activation(self):
        """Reset Windows activation status"""
        print("[Reset Windows Activation Status]")
        print()
        
        if not self.admin:
            print("[Error] This operation requires administrator privileges, please run the program as administrator")
            print()
            return False
        
        print("Warning: This operation will reset Windows activation status")
        print("This includes:")
        print("1. Uninstall current product key")
        print("2. Clear KMS server settings") 
        print("3. Reset authorization status")
        print("4. Clear activation cache")
        print()
        print("After reset, you will need to reactivate Windows")
        print()
        
        try:
            confirm = input("Confirm to reset Windows activation status? (y/N): ").strip().lower()
        except:
            confirm = 'n'
            
        if confirm != 'y':
            print("Reset cancelled")
            return False
        
        print()
        print("Resetting Windows activation status...")
        
        # 1. Uninstall product key
        print("1. Uninstalling product key...")
        result = self.run_slmgr_command("/upk", False)
        if "success" in result or "successfully" in result.lower():
            print("  [Success] Product key uninstalled")
        else:
            print("  [Warning] Error may have occurred while uninstalling product key")
        
        # 2. Clear KMS server settings
        print("2. Clearing KMS server settings...")
        result = self.run_slmgr_command("/ckms", False)
        if "success" in result or "successfully" in result.lower():
            print("  [Success] KMS server settings cleared")
        else:
            print("  [Warning] Error may have occurred while clearing KMS server settings")
        
        # 3. Clear product key from registry
        print("3. Clearing product key from registry...")
        result = self.run_slmgr_command("/cpky", False)
        if "success" in result or "successfully" in result.lower():
            print("  [Success] Product key cleared from registry")
        else:
            print("  [Warning] Error may have occurred while clearing product key from registry")
        
        # 4. Reset authorization status
        print("4. Resetting authorization status...")
        result = self.run_slmgr_command("/rearm", False)
        if "success" in result or "successfully" in result.lower():
            print("  [Success] Authorization status reset")
        else:
            print("  [Warning] Error may have occurred while resetting authorization status")
        
        print()
        print("[Success] Windows activation status reset")
        print("Note: You need to reactivate Windows to use all features")
        print("Recommend restarting computer to ensure all changes take effect")
        
        return True
    
    def show_reset_options(self):
        """Display reset options menu"""
        print("[Reset Options]")
        print()
        print("Please select reset option:")
        print("1. Complete activation status reset (recommended)")
        print("2. Uninstall product key only")
        print("3. Clear KMS server settings only")
        print("4. Reset authorization status only")
        print("5. Cancel")
        print()
        
        try:
            choice = input("Please select (1-5): ").strip()
        except:
            choice = '5'
        
        if choice == '1':
            return self.reset_windows_activation()
        elif choice == '2':
            return self.reset_product_key_only()
        elif choice == '3':
            return self.reset_kms_only()
        elif choice == '4':
            return self.reset_license_only()
        else:
            print("Reset cancelled")
            return False
    
    def reset_product_key_only(self):
        """Uninstall product key only"""
        print("[Uninstall Product Key Only]")
        
        if not self.admin:
            print("[Error] This operation requires administrator privileges")
            return False
        
        print("Uninstalling product key...")
        result = self.run_slmgr_command("/upk", False)
        
        if "success" in result or "successfully" in result.lower():
            print("[Success] Product key uninstalled")
            return True
        else:
            print(f"[Error] Failed to uninstall product key: {result}")
            return False
    
    def reset_kms_only(self):
        """Clear KMS server settings only"""
        print("[Clear KMS Server Settings Only]")
        
        if not self.admin:
            print("[Error] This operation requires administrator privileges")
            return False
        
        print("Clearing KMS server settings...")
        result = self.run_slmgr_command("/ckms", False)
        
        if "success" in result or "successfully" in result.lower():
            print("[Success] KMS server settings cleared")
            return True
        else:
            print(f"[Error] Failed to clear KMS server settings: {result}")
            return False
    
    def reset_license_only(self):
        """Reset authorization status only"""
        print("[Reset Authorization Status Only]")
        
        if not self.admin:
            print("[Error] This operation requires administrator privileges")
            return False
        
        print("Resetting authorization status...")
        result = self.run_slmgr_command("/rearm", False)
        
        if "success" in result or "successfully" in result.lower():
            print("[Success] Authorization status reset")
            print("Note: This operation consumes one reset count, usually has limit")
            return True
        else:
            print(f"[Error] Failed to reset authorization status: {result}")
            return False
    
    def show_menu(self):
        """Display menu"""
        print("Please select operation:")
        print("1. Check activation status")
        print("2. Display product keys")
        print("3. Install product key")
        print("4. Activate Windows")
        print("5. KMS activation")
        print("6. Query OEM information")
        print("7. Reset activation status")
        print("8. Exit program")
        print()
    
    def wait_for_enter(self):
        """Wait for user to press Enter"""
        print()
        try:
            input("Press Enter to continue...")
        except:
            print("Input error, continuing execution...")
            time.sleep(2)
    
    def main_loop(self):
        """Main loop"""
        while self.running:
            try:
                self.print_header()
                self.show_menu()
                
                choice = input("Please enter option (1-8): ").strip()
                
                if choice == '1':
                    self.print_header()
                    self.check_activation_status()
                    self.wait_for_enter()
                elif choice == '2':
                    self.print_header()
                    self.show_installed_product_keys()
                    self.wait_for_enter()
                elif choice == '3':
                    self.print_header()
                    if self.install_product_key():
                        time.sleep(2)
                        self.print_header()
                        self.check_activation_status()
                    self.wait_for_enter()
                elif choice == '4':
                    self.print_header()
                    if self.activate_windows():
                        time.sleep(2)
                        self.print_header()
                        self.check_activation_status()
                    self.wait_for_enter()
                elif choice == '5':
                    self.print_header()
                    if self.kms_activation():
                        time.sleep(2)
                        self.print_header()
                        self.check_activation_status()
                    self.wait_for_enter()
                elif choice == '6':
                    self.print_header()
                    self.get_oem_info()
                    self.wait_for_enter()
                elif choice == '7':
                    self.print_header()
                    if self.show_reset_options():
                        time.sleep(2)
                        self.print_header()
                        self.check_activation_status()
                    self.wait_for_enter()
                elif choice == '8':
                    print("Pressing Enter will exit automatically")
                    self.running = False
                else:
                    print("Invalid option, please re-enter")
                    time.sleep(1)
            except KeyboardInterrupt:
                print("\n\nProgram interrupted by user")
                self.running = False
            except Exception as e:
                print(f"\nProgram execution error: {str(e)}")
                self.wait_for_enter()

def main():
    """Main function"""
    try:
        manager = WindowsActivationManager()
        manager.main_loop()
    except Exception as e:
        print(f"Program execution error: {str(e)}")
    finally:
        try:
            input("\nPress Enter to exit...")
        except:
            pass

if __name__ == "__main__":
    main()