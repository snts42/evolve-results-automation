import json
import os
import base64
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from .logging_utils import log


class SecureCredentialManager:
    
    def __init__(self, encrypted_file: str):
        """Initialize with path to encrypted credentials file.
        
        Args:
            encrypted_file: Path to the encrypted credentials file (should have .enc extension)
        """
        self.encrypted_file = encrypted_file
        
    def _derive_key(self, password: str, salt: bytes) -> bytes:
        """Derive encryption key from password.
        
        Args:
            password: Master password
            salt: Random salt value
            
        Returns:
            bytes: Derived encryption key
        """
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=100000,
        )
        return base64.urlsafe_b64encode(kdf.derive(password.encode()))
    
    def decrypt_credentials(self, master_password: str) -> list:
        """Decrypt and return credentials from the encrypted file.
        
        Args:
            master_password: The master password for decryption
            
        Returns:
            list: Decrypted credentials
            
        Raises:
            FileNotFoundError: If the encrypted credentials file doesn't exist
            ValueError: If decryption fails (e.g., wrong password)
        """
        if not os.path.exists(self.encrypted_file):
            raise FileNotFoundError(f"No encrypted credentials file found: {self.encrypted_file}")
            
        try:
            with open(self.encrypted_file, 'rb') as f:
                data = f.read()
            
            if len(data) < 16:
                raise ValueError("Invalid encrypted file format")
                
            # Extract salt (first 16 bytes) and encrypted data
            salt = data[:16]
            encrypted_data = data[16:]
            
            # Derive key and decrypt
            key = self._derive_key(master_password, salt)
            fernet = Fernet(key)
            decrypted_data = fernet.decrypt(encrypted_data)
            
            # Parse and normalize credentials
            credentials = json.loads(decrypted_data.decode())
            if isinstance(credentials, dict):
                credentials = [credentials]
                
            return credentials
            
        except Exception as e:
            raise ValueError(f"Failed to decrypt credentials. Check password. Error: {e}")
    

    def remove_credential(self, username: str, master_password: str) -> bool:
        """Remove a credential by username from the encrypted credentials file.
        
        Args:
            username: The username of the credential to remove
            master_password: The master password for decryption/encryption
            
        Returns:
            bool: True if credential was removed, False if not found
            
        Raises:
            ValueError: If decryption fails or file operations fail
        """
        try:
            # Get current credentials
            credentials = self.decrypt_credentials(master_password)
            
            # Filter out the credential to remove
            initial_count = len(credentials)
            filtered = [cred for cred in credentials 
                      if cred.get('username') != username]
            
            # Check if anything was actually removed
            if len(filtered) == initial_count:
                log(f"Credential for username '{username}' not found.")
                return False
            
            # Save the updated credentials
            self._save_credentials(filtered, master_password)
            log(f"Credential for '{username}' removed.")
            return True
            
        except Exception as e:
            log(f"Error removing credential: {e}")
            return False

    def add_credential(self, username: str, password: str, master_password: str) -> bool:
        """Add a new credential to the encrypted credentials file.
        
        Args:
            username: The username to add
            password: The password for the username
            master_password: The master password for encryption
            
        Returns:
            bool: True if credential was added, False if username already exists
            
        Raises:
            ValueError: If encryption fails or file operations fail
        """
        try:
            # Get current credentials
            try:
                credentials = self.decrypt_credentials(master_password)
            except (FileNotFoundError, json.JSONDecodeError):
                # Handle case where file is empty or doesn't exist yet
                credentials = []
            
            # Check for duplicate username
            if any(cred.get('username') == username for cred in credentials):
                log(f"Credential for username '{username}' already exists.")
                return False
                
            # Add new credential and save
            credentials.append({'username': username, 'password': password})
            self._save_credentials(credentials, master_password)
            log(f"Credential for '{username}' added.")
            return True
            
        except Exception as e:
            log(f"Error adding credential: {e}")
            return False
            
    def _save_credentials(self, credentials: list, master_password: str) -> None:
        """Save credentials to the encrypted file.
        
        Args:
            credentials: List of credential dictionaries
            master_password: The master password for encryption
            
        Raises:
            ValueError: If encryption or file operations fail
        """
        salt = os.urandom(16)
        key = self._derive_key(master_password, salt)
        fernet = Fernet(key)
        
        # Convert credentials to JSON and encrypt
        encrypted_data = fernet.encrypt(json.dumps(credentials).encode())
        
        # Write salt + encrypted data to file
        with open(self.encrypted_file, 'wb') as f:
            f.write(salt + encrypted_data)

    def list_credentials(self, master_password: str = None) -> list:
        """Get all usernames from encrypted file. Returns list of credentials or empty list on error."""
        try:
            credentials = self.decrypt_credentials(master_password)
            return credentials
        except Exception as e:
            log(f"Error listing credentials: {e}")
            return []


def load_secure_credentials(credentials_file: str, master_password: str = None) -> list:
    """
    Load credentials securely. Supports both encrypted and plain text files.
    
    Args:
        credentials_file: Path to credentials file
        master_password: Master password for encrypted files (will prompt if not provided)
    
    Returns:
        List of credential dictionaries
    """
    manager = SecureCredentialManager(credentials_file)
    return manager.decrypt_credentials(master_password)