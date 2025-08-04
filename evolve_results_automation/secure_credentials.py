"""
Secure credential management with encryption support.
"""
import json
import os
import base64
import getpass
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from .logging_utils import log


class SecureCredentialManager:
    """Manages encrypted credential storage and retrieval."""
    
    def __init__(self, credentials_file: str):
        self.credentials_file = credentials_file
        self.encrypted_file = credentials_file.replace('.json', '.enc')
        
    def _derive_key(self, password: str, salt: bytes) -> bytes:
        """Derive encryption key from password."""
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=100000,
        )
        return base64.urlsafe_b64encode(kdf.derive(password.encode()))
    
    
        """Encrypt existing plain text credentials."""
        if not os.path.exists(self.credentials_file):
            log(f"Credentials file {self.credentials_file} not found")
            return False
            
        if not master_password:
            master_password = getpass.getpass("Enter master password to encrypt credentials: ")
            
        try:
            # Load plain text credentials
            with open(self.credentials_file, 'r') as f:
                credentials = json.load(f)
            
            # Generate salt and derive key
            salt = os.urandom(16)
            key = self._derive_key(master_password, salt)
            fernet = Fernet(key)
            
            # Encrypt credentials
            encrypted_data = fernet.encrypt(json.dumps(credentials).encode())
            
            # Save encrypted file with salt
            with open(self.encrypted_file, 'wb') as f:
                f.write(salt + encrypted_data)
            
            log(f"Credentials encrypted and saved to {self.encrypted_file}")
            
            # Optionally remove plain text file
            response = input("Delete plain text credentials file? [y/N]: ").strip().lower()
            if response == 'y':
                os.remove(self.credentials_file)
                log(f"Plain text credentials file {self.credentials_file} deleted")
            
            return True
            
        except Exception as e:
            log(f"Error encrypting credentials: {e}")
            return False
    
    def decrypt_credentials(self, master_password: str = None) -> list:
        """Decrypt and return credentials from the encrypted file."""
        if os.path.exists(self.encrypted_file):
            return self._load_encrypted_credentials(master_password)
        else:
            raise FileNotFoundError(f"No encrypted credentials file found: {self.encrypted_file}")
    
    def _load_encrypted_credentials(self, master_password: str = None) -> list:
        """Load encrypted credentials."""
        if not master_password:
            master_password = getpass.getpass("Enter master password: ")
            
        try:
            with open(self.encrypted_file, 'rb') as f:
                data = f.read()
            
            # Extract salt and encrypted data
            salt = data[:16]
            encrypted_data = data[16:]
            
            # Derive key and decrypt
            key = self._derive_key(master_password, salt)
            fernet = Fernet(key)
            decrypted_data = fernet.decrypt(encrypted_data)
            
            credentials = json.loads(decrypted_data.decode())
            if isinstance(credentials, dict):
                credentials = [credentials]
                
            return credentials
            
        except Exception as e:
            raise ValueError(f"Failed to decrypt credentials. Check password. Error: {e}")
    

    def remove_credential(self, username: str, master_password: str = None) -> bool:
        """Remove a credential by username from the encrypted credentials file."""
        try:
            credentials = self.decrypt_credentials(master_password)
            filtered = [cred for cred in credentials if cred.get('username') != username]
            if len(filtered) == len(credentials):
                log(f"Credential for username '{username}' not found.")
                return False
            # Re-encrypt and save
            if not master_password:
                master_password = getpass.getpass("Enter master password: ")
            salt = os.urandom(16)
            key = self._derive_key(master_password, salt)
            fernet = Fernet(key)
            encrypted_data = fernet.encrypt(json.dumps(filtered).encode())
            with open(self.encrypted_file, 'wb') as f:
                f.write(salt + encrypted_data)
            log(f"Credential for '{username}' removed.")
            return True
        except Exception as e:
            log(f"Error removing credential: {e}")
            return False

    def add_credential(self, username: str, password: str, master_password: str = None) -> bool:
        """Add a new credential to the encrypted credentials file."""
        try:
            credentials = self.decrypt_credentials(master_password)
            # Check for duplicate username
            if any(cred.get('username') == username for cred in credentials):
                log(f"Credential for username '{username}' already exists.")
                return False
            credentials.append({'username': username, 'password': password})
            # Re-encrypt and save
            if not master_password:
                master_password = getpass.getpass("Enter master password: ")
            # Generate new salt for each save
            salt = os.urandom(16)
            key = self._derive_key(master_password, salt)
            fernet = Fernet(key)
            encrypted_data = fernet.encrypt(json.dumps(credentials).encode())
            with open(self.encrypted_file, 'wb') as f:
                f.write(salt + encrypted_data)
            log(f"Credential for '{username}' added.")
            return True
        except Exception as e:
            log(f"Error adding credential: {e}")
            return False

    def list_credentials(self, master_password: str = None) -> bool:
        """List all usernames in encrypted file."""
        try:
            credentials = self.decrypt_credentials(master_password)
            print(f"\nFound {len(credentials)} credential(s):")
            for idx, cred in enumerate(credentials, 1):
                username = cred.get('username', '(missing username)')
                print(f"  [{idx}] {username}")
            return True
        except Exception as e:
            log(f"Error listing credentials: {e}")
            return False


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






    raise NotImplementedError("Plain text credential encryption is no longer supported.")
    """
    Interactive setup for credential encryption.
    
    Args:
        credentials_file: Path to plain text credentials file
        
    Returns:
        True if encryption was successful
    """
    manager = SecureCredentialManager(credentials_file)
    return manager.encrypt_credentials()
