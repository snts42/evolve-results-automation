"""
User Interface module for Evolve Results Automation.
Handles all user interactions, menus, and console output.
"""
import os
import sys
import getpass
from colorama import init, Fore, Style

from .config import ENCRYPTED_CREDENTIALS_FILE
from .secure_credentials import SecureCredentialManager, load_secure_credentials


class EvolveGUI:
    """Main GUI class for user interactions."""
    
    def __init__(self):
        init(autoreset=True)
        self.manager = SecureCredentialManager(ENCRYPTED_CREDENTIALS_FILE)
    
    def show_banner(self):
        """Display the application banner."""
        ascii_logo = f"""{Fore.GREEN}{Style.BRIGHT}
 _____ _   _ _____ _     _   _ _____    ___  _   _ _____ ________  ___ ___ _____ _____ _____ _   _ 
|  ___| | | |  _  | |   | | | |  ___|  / _ \| | | |_   _|  _  |  \/  |/ _ \_   _|_   _|  _  | \ | |
| |__ | | | | | | | |   | | | | |__   / /_\ \ | | | | | | | | | .  . / /_\ \| |   | | | | | |  \| |
|  __|| | | | | | | |   | | | |  __|  |  _  | | | | | | | | | | |\/| |  _  || |   | | | | | | . ` |
| |___\ \_/ | \_/ / |___\ \_/ / |___  | | | | |_| | | | \ \_/ / |  | | | | || |  _| |_\ \_/ / |\  |
\____/ \___/ \___/\_____/\___/\____/  \_| |_/\___/  \_/  \___/\_|  |_|_| |_/\_/  \___/ \___/\_| \_/
{Style.RESET_ALL}
{Fore.WHITE}{Style.BRIGHT}
                  Evolve Results Automation Tool by snts42
{Style.RESET_ALL}
"""
        print(ascii_logo)
        print(Fore.GREEN + "=" * 100)
        print(Fore.WHITE + "WELCOME TO THE EVOLVE RESULTS AUTOMATION TOOL!")
        print(Fore.GREEN + "=" * 100 + Style.RESET_ALL)
        print(
            Fore.GREEN +
            "\nAutomate the retrieval and download of exam results from the City & Guilds Evolve platform."
        )
    
    def show_main_menu(self, master_password=None):
        """Display main menu and get user choice."""
        if master_password:
            self.master_password = master_password
        while True:
            print(f"\n{Fore.CYAN}{'='*60}")
            print(f"{Fore.CYAN}MAIN MENU")
            print(f"{Fore.CYAN}{'='*60}")
            print(f"{Fore.WHITE}1. Run Results Automation")
            print(f"{Fore.WHITE}2. Manage Credentials")
            print(f"{Fore.WHITE}3. Exit")
            
            choice = input(f"\n{Fore.GREEN}Select option [1-3]: {Style.RESET_ALL}").strip()
            
            if choice == '1':
                return 'run_automation'
            elif choice == '2':
                self.show_credentials_menu(master_password=self.master_password)
            elif choice == '3':
                print(f"{Fore.CYAN}Goodbye!{Style.RESET_ALL}")
                sys.exit(0)
            else:
                print(f"{Fore.RED}Invalid option. Please select 1-3.{Style.RESET_ALL}")
    
    def setup_automation(self, master_password=None):
        """Setup for running the main automation."""
        if master_password:
            self.master_password = master_password

        
        # Check credentials exist
        if not self._credentials_exist():
            print(Fore.RED + f"\nERROR: No credentials found!")
            print(Fore.YELLOW + "Please use 'Manage Credentials' option to set up your accounts first.")
            input(Fore.CYAN + "Press Enter to return to main menu..." + Style.RESET_ALL)
            return None
        
        # Load and display credentials
        try:
            accounts = load_secure_credentials(ENCRYPTED_CREDENTIALS_FILE, master_password=self.master_password)
            self._display_credentials(accounts)
        except Exception as e:
            print(Fore.RED + f"\nERROR: Unable to load credentials: {e}")
            input(Fore.CYAN + "Press Enter to return to main menu..." + Style.RESET_ALL)
            return None
        # Get headless preference
        return self._get_headless_preference()
    
    def show_credentials_menu(self, master_password=None):
        """Display credential management menu."""
        if master_password:
            self.master_password = master_password
        # Onboarding: if credentials.enc missing, prompt to set master password before menu
        enc_file = ENCRYPTED_CREDENTIALS_FILE
        if not os.path.exists(enc_file):
            print(f"{Fore.YELLOW}\nNo encrypted credentials file found. You must set a master password to begin.")
            while True:
                pw1 = getpass.getpass(f"{Fore.CYAN}Set a new master password: {Style.RESET_ALL}").strip()
                pw2 = getpass.getpass(f"{Fore.CYAN}Confirm master password: {Style.RESET_ALL}").strip()
                if not pw1:
                    print(f"{Fore.RED}❌ Master password cannot be empty.{Style.RESET_ALL}")
                elif pw1 != pw2:
                    print(f"{Fore.RED}❌ Passwords do not match. Please try again.{Style.RESET_ALL}")
                else:
                    self.master_password = pw1
                    print(f"{Fore.GREEN}Master password set. Add your first credential to create the encrypted file.{Style.RESET_ALL}")
                    break
        while True:
            print(f"\n{Fore.CYAN}{'='*60}")
            print(f"{Fore.CYAN}CREDENTIAL MANAGEMENT")
            print(f"{Fore.CYAN}{'='*60}")
            print(f"{Fore.WHITE}1. List credentials")
            print(f"{Fore.WHITE}2. Add credential")
            print(f"{Fore.WHITE}3. Remove credential")
            print(f"{Fore.WHITE}4. Back to main menu")
            
            choice = input(f"\n{Fore.GREEN}Select option [1-4]: {Style.RESET_ALL}").strip()
            
            if choice == '1':
                self._handle_list_credentials()
            elif choice == '2':
                self._handle_add_credential()
            elif choice == '3':
                self._handle_remove_credential()
            elif choice == '4':
                break
            else:
                print(f"{Fore.RED}❌ Invalid option. Please select 1-4.{Style.RESET_ALL}")
            
            input(f"\n{Fore.CYAN}Press Enter to continue...{Style.RESET_ALL}")
    
    def show_progress(self, current: int, total: int, message: str):
        """Display progress information."""
        percentage = (current / total) * 100 if total > 0 else 0
        print(f"{Fore.CYAN}[{current}/{total}] {percentage:.1f}% - {message}{Style.RESET_ALL}")
    
    def show_summary(self, stats):
        """Display execution summary."""
        print(f"\n{Fore.GREEN}{'='*60}")
        print(f"{Fore.GREEN}🎉 AUTOMATION COMPLETED!")
        print(f"{Fore.GREEN}{'='*60}")
        print(f"{Fore.WHITE}📊 Accounts processed: {stats.accounts_processed}")
        print(f"{Fore.WHITE}📝 New rows added: {stats.new_rows_added}")
        print(f"{Fore.WHITE}📄 PDFs downloaded: {stats.pdfs_downloaded}")
        print(f"{Fore.WHITE}❌ Errors encountered: {stats.errors_encountered}")
        print(f"{Fore.GREEN}{'='*60}{Style.RESET_ALL}")
    
    # Private helper methods
    def _credentials_exist(self) -> bool:
        """Check if any credentials file exists."""
        encrypted_file = ENCRYPTED_CREDENTIALS_FILE.replace('.json', '.enc')
        return os.path.exists(ENCRYPTED_CREDENTIALS_FILE) or os.path.exists(encrypted_file)
    
        encrypted_file = ENCRYPTED_CREDENTIALS_FILE.replace('.json', '.enc')
        return os.path.exists(ENCRYPTED_CREDENTIALS_FILE) and not os.path.exists(encrypted_file)
    

        """Offer to encrypt plain text credentials."""
        print(Fore.YELLOW + "\n⚠️  WARNING: Credentials are stored in plain text!")
        encrypt_choice = input(Fore.CYAN + "Would you like to encrypt them now for security? [Y/n]: " + Style.RESET_ALL).strip().lower()
        if encrypt_choice in ("", "y", "yes"):
            if setup_credential_encryption(ENCRYPTED_CREDENTIALS_FILE):
                print(Fore.GREEN + "✅ Credentials encrypted successfully!")
                return True
            else:
                print(Fore.RED + "❌ Failed to encrypt credentials.")
                return False
        return True
    
    def _display_credentials(self, accounts: list):
        """Display loaded credentials."""
        print(Fore.GREEN + f"\nFound {len(accounts)} credential(s):")
        for idx, acc in enumerate(accounts, 1):
            user = acc.get('username', '(missing username)')
            print(Fore.CYAN + f"  [{idx}] {user}")
    
    def _get_headless_preference(self) -> bool:
        """Get user preference for headless mode."""
        while True:
            headless_input = input(
                Fore.GREEN + "\nRun in HEADLESS (no browser window) mode? [Y/n]: " + Style.RESET_ALL
            ).strip().lower()
            if headless_input in ("", "y", "yes"):
                return True
            elif headless_input in ("n", "no"):
                return False
            else:
                print(Fore.YELLOW + "Please enter 'Y' or 'N'." + Style.RESET_ALL)
    
    def _handle_list_credentials(self):
        """Handle listing credentials."""
        enc_file = ENCRYPTED_CREDENTIALS_FILE
        if not os.path.exists(enc_file):
            print(f"{Fore.YELLOW}No credentials saved yet. Add your first credential to create the encrypted file.{Style.RESET_ALL}")
            return
        if not hasattr(self, 'master_password') or not self.master_password:
            self.master_password = getpass.getpass(f"{Fore.CYAN}Enter master password for encrypted credentials: {Style.RESET_ALL}").strip()
            if not self.master_password:
                print(f"{Fore.RED}❌ Master password cannot be empty.{Style.RESET_ALL}")
                return
        try:
            self.manager.list_credentials(master_password=self.master_password)
        except Exception as e:
            print(f"{Fore.RED}❌ Error: {e}{Style.RESET_ALL}")
    
    def _handle_add_credential(self):
        """Handle adding a credential."""
        username = input(f"{Fore.CYAN}Enter username: {Style.RESET_ALL}").strip()
        password = getpass.getpass(f"{Fore.CYAN}Enter password: {Style.RESET_ALL}").strip()

        # Prompt for master password if not set (onboarding or after reset)
        enc_file = ENCRYPTED_CREDENTIALS_FILE
        if not hasattr(self, 'master_password') or not self.master_password:
            if not os.path.exists(enc_file):
                # Already set in show_credentials_menu, but double-check for robustness
                self.master_password = getpass.getpass(f"{Fore.CYAN}Set master password for encrypted credentials: {Style.RESET_ALL}").strip()
                if not self.master_password:
                    print(f"{Fore.RED}❌ Master password cannot be empty.{Style.RESET_ALL}")
                    return
            else:
                self.master_password = getpass.getpass(f"{Fore.CYAN}Enter master password for encrypted credentials: {Style.RESET_ALL}").strip()
                if not self.master_password:
                    print(f"{Fore.RED}❌ Master password cannot be empty.{Style.RESET_ALL}")
                    return

        if username and password:
            try:
                # Ensure credentials.enc is initialized if missing
                if not os.path.exists(enc_file):
                    import json
                    salt = os.urandom(16)
                    from cryptography.fernet import Fernet
                    key = self.manager._derive_key(self.master_password, salt)
                    fernet = Fernet(key)
                    encrypted_data = fernet.encrypt(json.dumps([]).encode())
                    with open(enc_file, 'wb') as f:
                        f.write(salt + encrypted_data)
                if self.manager.add_credential(username, password, master_password=self.master_password):
                    print(f"{Fore.GREEN}✅ Credential added successfully!{Style.RESET_ALL}")
                else:
                    print(f"{Fore.RED}❌ Failed to add credential.{Style.RESET_ALL}")
            except Exception as e:
                print(f"{Fore.RED}❌ Error: {e}{Style.RESET_ALL}")
        else:
            print(f"{Fore.RED}❌ Username and password cannot be empty.{Style.RESET_ALL}")
    
    def _handle_remove_credential(self):
        """Handle removing a credential."""
        username = input(f"{Fore.CYAN}Enter username to remove: {Style.RESET_ALL}").strip()

        # Prompt for master password if not set (onboarding or after reset)
        if not hasattr(self, 'master_password') or not self.master_password:
            self.master_password = getpass.getpass(f"{Fore.CYAN}Enter master password for encrypted credentials: {Style.RESET_ALL}").strip()
            if not self.master_password:
                print(f"{Fore.RED}❌ Master password cannot be empty.{Style.RESET_ALL}")
                return

        if username:
            try:
                if self.manager.remove_credential(username, master_password=self.master_password):
                    print(f"{Fore.GREEN}✅ Credential removed successfully!{Style.RESET_ALL}")
                else:
                    print(f"{Fore.RED}❌ Failed to remove credential.{Style.RESET_ALL}")
            except Exception as e:
                print(f"{Fore.RED}❌ Error: {e}{Style.RESET_ALL}")
        else:
            print(f"{Fore.RED}❌ Username cannot be empty.{Style.RESET_ALL}")
    


# Convenience functions for backward compatibility
def show_banner():
    """Show application banner."""
    gui = EvolveGUI()
    gui.show_banner()

def show_main_menu():
    """Show main menu and return choice."""
    gui = EvolveGUI()
    return gui.show_main_menu()

def setup_automation():
    """Setup automation and return headless preference."""
    gui = EvolveGUI()
    return gui.setup_automation()
