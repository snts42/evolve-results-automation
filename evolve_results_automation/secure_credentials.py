import json
import os
import hmac
import hashlib
import logging
import tempfile
import pyaes


# Format: salt(16) + iv(16) + mac(32) + ciphertext(N)
_SALT_LEN = 16
_IV_LEN = 16
_MAC_LEN = 32
_HEADER_LEN = _SALT_LEN + _IV_LEN + _MAC_LEN


class SecureCredentialManager:
    
    def __init__(self, encrypted_file: str):
        """Initialize with path to encrypted credentials file.
        
        Args:
            encrypted_file: Path to the encrypted credentials file (should have .enc extension)
        """
        self.encrypted_file = encrypted_file
        
    def _derive_keys(self, password: str, salt: bytes) -> tuple:
        """Derive AES-256 encryption key and HMAC key from password.
        
        Args:
            password: Master password
            salt: Random salt value
            
        Returns:
            tuple: (aes_key: 32 bytes, hmac_key: 32 bytes)
        """
        # Derive 64 bytes: first 32 for AES-256, last 32 for HMAC-SHA256
        derived = hashlib.pbkdf2_hmac('sha256', password.encode(), salt, 100000, dklen=64)
        return derived[:32], derived[32:]

    def _encrypt(self, plaintext: bytes, password: str) -> bytes:
        """Encrypt plaintext with AES-256-CBC + HMAC-SHA256.
        
        Returns:
            bytes: salt(16) + iv(16) + hmac(32) + ciphertext
        """
        salt = os.urandom(_SALT_LEN)
        iv = os.urandom(_IV_LEN)
        aes_key, hmac_key = self._derive_keys(password, salt)

        # PKCS7 pad and AES-256-CBC encrypt
        encrypter = pyaes.Encrypter(pyaes.AESModeOfOperationCBC(aes_key, iv=iv))
        ciphertext = encrypter.feed(plaintext) + encrypter.feed()

        # Authenticate: HMAC-SHA256 over iv + ciphertext
        mac = hmac.new(hmac_key, iv + ciphertext, hashlib.sha256).digest()

        return salt + iv + mac + ciphertext

    def _decrypt(self, data: bytes, password: str) -> bytes:
        """Decrypt AES-256-CBC + HMAC-SHA256 data.
        
        Returns:
            bytes: Decrypted plaintext
            
        Raises:
            ValueError: If MAC verification fails or data is malformed
        """
        if len(data) < _HEADER_LEN + 16:  # minimum: header + 1 AES block
            raise ValueError("Invalid encrypted file format")

        salt = data[:_SALT_LEN]
        iv = data[_SALT_LEN:_SALT_LEN + _IV_LEN]
        stored_mac = data[_SALT_LEN + _IV_LEN:_HEADER_LEN]
        ciphertext = data[_HEADER_LEN:]

        aes_key, hmac_key = self._derive_keys(password, salt)

        # Verify HMAC before decrypting (authenticate-then-decrypt)
        computed_mac = hmac.new(hmac_key, iv + ciphertext, hashlib.sha256).digest()
        if not hmac.compare_digest(stored_mac, computed_mac):
            raise ValueError("HMAC verification failed - wrong password or corrupted file")

        # AES-256-CBC decrypt with PKCS7 unpadding
        decrypter = pyaes.Decrypter(pyaes.AESModeOfOperationCBC(aes_key, iv=iv))
        plaintext = decrypter.feed(ciphertext) + decrypter.feed()

        return plaintext
    
    def create_empty(self, master_password: str) -> None:
        """Create an empty encrypted credentials file with the given master password.
        
        Args:
            master_password: The master password for encryption
        """
        self._save_credentials([], master_password)

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
            
            decrypted_data = self._decrypt(data, master_password)
            
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
            try:
                credentials = self.decrypt_credentials(master_password)
            except (FileNotFoundError, ValueError):
                # Handle case where file doesn't exist or is corrupted
                credentials = []
            
            # Filter out the credential to remove
            initial_count = len(credentials)
            filtered = [cred for cred in credentials 
                      if cred.get('username') != username]
            
            # Check if anything was actually removed
            if len(filtered) == initial_count:
                logging.info(f"Credential for username '{username}' not found")
                return False
            
            # Save the updated credentials
            self._save_credentials(filtered, master_password)
            logging.info(f"Credential for '{username}' removed")
            return True
            
        except Exception as e:
            logging.error(f"Error removing credential: {e}")
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
            except (FileNotFoundError, ValueError):
                # Handle case where file is empty or doesn't exist yet
                credentials = []
            
            # Check for duplicate username
            if any(cred.get('username') == username for cred in credentials):
                logging.info(f"Credential for username '{username}' already exists")
                return False
                
            # Add new credential and save
            credentials.append({'username': username, 'password': password})
            self._save_credentials(credentials, master_password)
            logging.info(f"Credential for '{username}' added")
            return True
            
        except Exception as e:
            logging.error(f"Error adding credential: {e}")
            return False
            
    def _save_credentials(self, credentials: list, master_password: str) -> None:
        """Save credentials to the encrypted file.
        
        Args:
            credentials: List of credential dictionaries
            master_password: The master password for encryption
            
        Raises:
            ValueError: If encryption or file operations fail
        """
        plaintext = json.dumps(credentials).encode()
        encrypted_data = self._encrypt(plaintext, master_password)
        
        # Atomic write: write to temp file then replace to prevent corruption
        dir_name = os.path.dirname(self.encrypted_file) or '.'
        fd, tmp_path = tempfile.mkstemp(dir=dir_name)
        try:
            os.write(fd, encrypted_data)
            os.close(fd)
            os.replace(tmp_path, self.encrypted_file)
        except Exception as e:
            try:
                os.close(fd)
            except OSError:
                pass
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
            raise

    def list_credentials(self, master_password: str = None) -> list:
        """Get all usernames from encrypted file. Returns list of usernames or empty list on error."""
        try:
            credentials = self.decrypt_credentials(master_password)
            return [{"username": cred.get("username", "?")} for cred in credentials]
        except Exception as e:
            logging.error(f"Error listing credentials: {e}")
            return []