import os
import sys
import getpass
from colorama import init, Fore, Style

from .config import ENCRYPTED_CREDENTIALS_FILE
from .secure_credentials import SecureCredentialManager, load_secure_credentials


class EvolveGUI:
    
    def __init__(self):
        init(autoreset=True)
        self.manager = SecureCredentialManager(ENCRYPTED_CREDENTIALS_FILE)
    
    def show_banner(self):
        ascii_art = r""" _____ _   _ _____ _     _   _ _____    ___  _   _ _____ ________  ___ ___ _____ _____ _____ _   _ 
|  ___| | | |  _  | |   | | | |  ___|  / _ \| | | |_   _|  _  |  \/  |/ _ \_   _|_   _|  _  | \ | |
| |__ | | | | | | | |   | | | | |__   / /_\ \ | | | | | | | | | .  . / /_\ \| |   | | | | | |  \| |
|  __|| | | | | | | |   | | | |  __|  |  _  | | | | | | | | | | |\/| |  _  || |   | | | | | | . ` |
| |___\ \_/ | \_/ / |___\ \_/ / |___  | | | | |_| | | | \ \_/ / |  | | | | || |  _| |_\ \_/ / |\  |
\____/ \___/ \___/\_____/\___/\____/  \_| |_/\___/  \_/  \___/\_|  |_|_| |_/\_/  \___/ \___/\_| \_/"""
        
        print(f"{Fore.GREEN}{ascii_art}{Style.RESET_ALL}")
        print(f"{Fore.WHITE + Style.BRIGHT}                            Evolve Results Automation Tool by snts42{Style.RESET_ALL}")
        print()
        print(f"{Fore.GREEN}{'='*100}{Style.RESET_ALL}")
        print(f"{Fore.GREEN + Style.BRIGHT}                          Welcome to Evolve Results Automation Tool{Style.RESET_ALL}")
        print(f"{Fore.GREEN}{'='*100}{Style.RESET_ALL}")
        print(f"{Fore.YELLOW + Style.BRIGHT}\n IMPORTANT:{Style.RESET_ALL} {Fore.YELLOW}If you forget your master password, delete 'credentials.enc' and restart the program.{Style.RESET_ALL}")
        print(f"{Fore.YELLOW}   This will {Fore.RED + Style.BRIGHT}PERMANENTLY ERASE{Style.RESET_ALL} {Fore.YELLOW}all saved credentials.{Style.RESET_ALL}")
        print(f"{Fore.GREEN}{'='*100}{Style.RESET_ALL}")
    
    def show_main_menu(self, master_password=None):
        if master_password:
            self.master_password = master_password
        while True:
            print(f"\n{Fore.GREEN}{'='*100}{Style.RESET_ALL}")
            print(f"{Fore.GREEN + Style.BRIGHT}{'MAIN MENU'.center(100)}{Style.RESET_ALL}")
            print(f"{Fore.GREEN}{'='*100}{Style.RESET_ALL}")
            print()
            print(f"{Fore.GREEN}    * 1. Run Results Automation{Style.RESET_ALL}")
            print(f"{Fore.GREEN}    * 2. Manage Credentials{Style.RESET_ALL}")
            print(f"{Fore.GREEN}    * 3. Exit{Style.RESET_ALL}")
            print()
            print(f"{Fore.GREEN}{'='*100}{Style.RESET_ALL}")
            print()
            choice = input(f"\n{Fore.WHITE + Style.BRIGHT}➤ Select option [1-3]: {Style.RESET_ALL}").strip()
            
            if choice == '1':
                return 'run_automation'
            elif choice == '2':
                self.show_credentials_menu(master_password=self.master_password)
            elif choice == '3':
                print(f"\n{Fore.GREEN + Style.BRIGHT}GOODBYE!{Style.RESET_ALL}")
                sys.exit(0)
            else:
                print(f"{Fore.RED}Invalid option. Please select 1-3.{Style.RESET_ALL}")
    
    def setup_automation(self, master_password=None):
        """Setup for running the main automation."""
        if master_password:
            self.master_password = master_password

        
        # Check credentials exist
        if not os.path.exists(ENCRYPTED_CREDENTIALS_FILE):
            print(f"{Fore.RED + Style.BRIGHT}\nERROR: No Evolve credentials found!{Style.RESET_ALL}")
            print(f"{Fore.YELLOW + Style.BRIGHT}Please use 'Manage Credentials' option to set up your accounts first.{Style.RESET_ALL}")
            input(f"{Fore.WHITE + Style.BRIGHT}Press Enter to return to main menu...{Style.RESET_ALL}")
            return None
        
        # Load and display credentials
        try:
            accounts = load_secure_credentials(ENCRYPTED_CREDENTIALS_FILE, master_password=self.master_password)
            self._display_credentials(accounts)
        except Exception as e:
            print(Fore.RED + f"\nERROR: Unable to load credentials: {e}")
            input(Fore.WHITE + Style.BRIGHT + "Press Enter to return to main menu..." + Style.RESET_ALL)
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
            print(f"{Fore.YELLOW}\nNo Evolve credentials found. You must set a master password to begin.")
            while True:
                pw1 = getpass.getpass(f"{Fore.WHITE}Set a new master password: {Style.RESET_ALL}").strip()
                pw2 = getpass.getpass(f"{Fore.WHITE}Confirm master password: {Style.RESET_ALL}").strip()
                if not pw1:
                    print(f"{Fore.RED}Master password cannot be empty.{Style.RESET_ALL}")
                elif pw1 != pw2:
                    print(f"{Fore.RED}Passwords do not match. Please try again.{Style.RESET_ALL}")
                else:
                    self.master_password = pw1
                    print(f"{Fore.GREEN}Master password set. Add your first credential to create the encrypted file.{Style.RESET_ALL}")
                    break
        while True:
            print(f"\n{Fore.GREEN}{'='*100}{Style.RESET_ALL}")
            print(f"{Fore.GREEN + Style.BRIGHT}{'CREDENTIAL MANAGEMENT'.center(100)}{Style.RESET_ALL}")
            print(f"{Fore.GREEN}{'='*100}{Style.RESET_ALL}")
            print()
            print(f"{Fore.GREEN}    * 1. List credentials{Style.RESET_ALL}")
            print(f"{Fore.GREEN}    * 2. Add credential{Style.RESET_ALL}")
            print(f"{Fore.GREEN}    * 3. Remove credential{Style.RESET_ALL}")
            print(f"{Fore.GREEN}    * 4. Back to main menu{Style.RESET_ALL}")
            print()
            print(f"{Fore.GREEN}{'='*100}{Style.RESET_ALL}")
            print()
            choice = input(f"\n{Fore.WHITE + Style.BRIGHT}➤ Select option [1-4]: {Style.RESET_ALL}").strip()
            
            if choice == '1':
                self._handle_list_credentials()
            elif choice == '2':
                self._handle_add_credential()
            elif choice == '3':
                self._handle_remove_credential()
            elif choice == '4':
                break
            else:
                print(f"{Fore.RED}Invalid option. Please select 1-4.{Style.RESET_ALL}")
            
            if choice in ['1', '2', '3']:
                input(f"\n{Fore.WHITE}Press Enter to continue...{Style.RESET_ALL}")
    
    def show_progress(self, current: int, total: int, message: str):
        """Display progress information."""
        percentage = (current / total) * 100 if total > 0 else 0
        print(f"{Fore.GREEN}[{current:>3}/{total}] {Fore.YELLOW}{percentage:>5.1f}%{Style.RESET_ALL} {Fore.WHITE}- {message}{Style.RESET_ALL}")
    
    def show_summary(self, stats):
        """Display execution summary."""
        print(f"\n{Fore.GREEN}{'='*100}{Style.RESET_ALL}")
        print(f"{Fore.GREEN + Style.BRIGHT}                              AUTOMATION COMPLETED!{Style.RESET_ALL}")
        print(f"{Fore.GREEN}{'='*100}{Style.RESET_ALL}")
        print(f"{Fore.WHITE}Accounts processed: {stats.accounts_processed}")
        print(f"{Fore.WHITE}New rows added: {stats.new_rows_added}")
        print(f"{Fore.WHITE}PDFs downloaded: {stats.pdfs_downloaded}")
        print(f"{Fore.WHITE}Errors encountered: {stats.errors_encountered}")
        print(f"{Fore.GREEN}{'='*100}{Style.RESET_ALL}")

    def _display_credentials(self, accounts: list):
        """Display loaded credentials."""
        print(f"\n{Fore.GREEN + Style.BRIGHT}Found {len(accounts)} Evolve credential(s):{Style.RESET_ALL}")
        for idx, acc in enumerate(accounts, 1):
            user = acc.get('username', '(missing username)')
            print(f"{Fore.GREEN}  [{idx}] {user}{Style.RESET_ALL}")
    
    def _get_headless_preference(self) -> bool:
        """Get user preference for headless mode."""
        while True:
            headless_input = input(
                f"{Fore.WHITE + Style.BRIGHT}Run in HEADLESS mode? [Y/n]: {Style.RESET_ALL}"
            ).strip().lower()
            if headless_input in ("", "y", "yes"):
                return True
            elif headless_input in ("n", "no"):
                return False
            else:
                print(f"{Fore.YELLOW}Please enter 'Y' or 'N'.{Style.RESET_ALL}")
    
    def _handle_list_credentials(self):
        """Handle listing credentials."""
        enc_file = ENCRYPTED_CREDENTIALS_FILE
        if not os.path.exists(enc_file):
            print(f"{Fore.YELLOW}No credentials saved yet. Add your first credential to create the encrypted file.{Style.RESET_ALL}")
            return
        if not hasattr(self, 'master_password') or not self.master_password:
            self.master_password = getpass.getpass(f"{Fore.WHITE + Style.BRIGHT}Enter master password for encrypted credentials: {Style.RESET_ALL}").strip()
            if not self.master_password:
                print(f"{Fore.RED}X Master password cannot be empty.{Style.RESET_ALL}")
                return
        try:
            credentials = self.manager.list_credentials(master_password=self.master_password)
            if credentials:
                self._display_credentials(credentials)
            else:
                print(f"{Fore.YELLOW + Style.BRIGHT}No credentials found.{Style.RESET_ALL}")
        except Exception as e:
            print(f"{Fore.RED}Error: {e}{Style.RESET_ALL}")
    
    def _handle_add_credential(self):
        username = input(f"{Fore.WHITE}Enter username: {Style.RESET_ALL}").strip()
        password = getpass.getpass(f"{Fore.WHITE}Enter password: {Style.RESET_ALL}").strip()

        # Prompt for master password if not set (onboarding or after reset)
        enc_file = ENCRYPTED_CREDENTIALS_FILE
        if not hasattr(self, 'master_password') or not self.master_password:
            if not os.path.exists(enc_file):
                self.master_password = getpass.getpass(f"{Fore.WHITE}Set master password for encrypted credentials: {Style.RESET_ALL}").strip()
                if not self.master_password:
                    print(f"{Fore.RED}Master password cannot be empty.{Style.RESET_ALL}")
                    return
            else:
                self.master_password = getpass.getpass(f"{Fore.WHITE}Enter master password for encrypted credentials: {Style.RESET_ALL}").strip()
                if not self.master_password:
                    print(f"{Fore.RED}Master password cannot be empty.{Style.RESET_ALL}")
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
                    print(f"{Fore.GREEN} Credential added successfully!{Style.RESET_ALL}")
                else:
                    print(f"{Fore.RED}Failed to add credential.{Style.RESET_ALL}")
            except Exception as e:
                print(f"{Fore.RED}Error: {e}{Style.RESET_ALL}")
        else:
            print(f"{Fore.RED}Username and password cannot be empty.{Style.RESET_ALL}")
    
    def _handle_remove_credential(self):
        """Handle removing a credential."""
        username = input(f"{Fore.WHITE}Enter username to remove: {Style.RESET_ALL}").strip()

        # Prompt for master password if not set (onboarding or after reset)
        if not hasattr(self, 'master_password') or not self.master_password:
            self.master_password = getpass.getpass(f"{Fore.WHITE + Style.BRIGHT}Enter master password for encrypted credentials: {Style.RESET_ALL}").strip()
            if not self.master_password:
                print(f"{Fore.RED}Master password cannot be empty.{Style.RESET_ALL}")
                return

        if username:
            try:
                if self.manager.remove_credential(username, master_password=self.master_password):
                    print(f"{Fore.GREEN}Credential removed successfully!{Style.RESET_ALL}")
                else:
                    print(f"{Fore.RED}Failed to remove credential.{Style.RESET_ALL}")
            except Exception as e:
                print(f"{Fore.RED}Error: {e}{Style.RESET_ALL}")
        else:
            print(f"{Fore.RED}Username cannot be empty.{Style.RESET_ALL}")
    
    def handle_master_password_setup(self):
        """Handle master password setup and validation. Returns master password or None if cancelled."""
        if not os.path.exists(ENCRYPTED_CREDENTIALS_FILE):
            # Onboarding: No credentials file exists
            self.show_credentials_menu()
            # After adding, check again
            if not os.path.exists(ENCRYPTED_CREDENTIALS_FILE):
                print(f"{Fore.RED}No credentials were added. Exiting.{Style.RESET_ALL}")
                return None
        
        # Loop until valid master password is entered
        manager = SecureCredentialManager(ENCRYPTED_CREDENTIALS_FILE)
        while True:
            master_password = getpass.getpass(f"{Fore.WHITE}Enter master password: {Style.RESET_ALL}")
            try:
                # Try to decrypt credentials to validate password
                manager.decrypt_credentials(master_password)
                self.master_password = master_password
                return master_password
            except Exception:
                print(f"{Fore.RED}Invalid password.{Style.RESET_ALL}")
    
    def wait_for_continue(self):
        """Wait for user to press Enter to return to main menu."""
        input(f"\n{Fore.WHITE}Press Enter to return to main menu...{Style.RESET_ALL}")
    
    def show_cancellation_message(self):
        """Show operation cancelled message."""
        print(f"\n{Fore.YELLOW}Operation cancelled by user.{Style.RESET_ALL}")
    
    def show_fatal_error(self, error_message: str):
        """Show fatal error message."""
        print(f"\n{Fore.RED}Fatal error: {error_message}{Style.RESET_ALL}")
