import json

def load_credentials(json_file: str):
    with open(json_file, 'r') as f:
        creds = json.load(f)
    if isinstance(creds, dict):
        creds = [creds]
    return creds