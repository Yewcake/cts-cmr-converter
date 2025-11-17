"""
Auto-updater for CTS CMR Converter
Checks for updates on network share or GitHub
"""

import json
import os
from pathlib import Path
from typing import Optional, Dict
from packaging import version as ver

# Current version
__version__ = "1.0.3"

# Update sources (configure based on your setup)
UPDATE_CHECK_ENABLED = True

# Option 1: Network share (recommended for internal company use)
NETWORK_SHARE_PATH = r"\\cts-server\software\CMR_Converter\version.json"

# Option 2: GitHub (for public releases)
GITHUB_REPO = "yewcake/cts-cmr-converter"
GITHUB_TOKEN = ""

# Option 3: Web server
WEB_SERVER_URL = "https://yourserver.com/cmr_converter/version.json"


def check_for_updates_network_share() -> Optional[Dict]:
    """
    Check for updates on company network share
    
    Network share structure:
    \\cts-server\software\CMR_Converter\
    ├── version.json
    ├── CTS_CMR_Converter_Setup.exe (latest)
    └── updates\
        ├── CTS_CMR_Converter_Setup_v1.0.0.exe
        └── CTS_CMR_Converter_Setup_v1.0.1.exe
    
    version.json format:
    {
        "version": "1.0.1",
        "download_path": "\\\\cts-server\\software\\CMR_Converter\\CTS_CMR_Converter_Setup.exe",
        "release_notes": "Bug fixes:\n- Fixed address formatting\n- Improved box detection"
    }
    """
    try:
        if not os.path.exists(NETWORK_SHARE_PATH):
            return None
        
        with open(NETWORK_SHARE_PATH, 'r') as f:
            data = json.load(f)
        
        latest_version = data['version']
        
        # Compare versions
        if ver.parse(latest_version) > ver.parse(__version__):
            return {
                'available': True,
                'version': latest_version,
                'download_path': data['download_path'],
                'release_notes': data.get('release_notes', 'No release notes available'),
                'source': 'network_share'
            }
        
        return {'available': False}
    
    except Exception as e:
        print(f"Network share update check failed: {e}")
        return None


def check_for_updates_github() -> Optional[Dict]:
    """
    Check for updates on GitHub releases
    Requires: pip install requests
    """
    try:
        import requests
        
        url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        response = requests.get(url, timeout=5)
        
        if response.status_code == 200:
            data = response.json()
            latest_version = data['tag_name'].lstrip('v')
            
            if ver.parse(latest_version) > ver.parse(__version__):
                # Find the .exe asset
                download_url = None
                for asset in data.get('assets', []):
                    if asset['name'].endswith('.exe'):
                        download_url = asset['browser_download_url']
                        break
                
                if download_url:
                    return {
                        'available': True,
                        'version': latest_version,
                        'download_url': download_url,
                        'release_notes': data.get('body', 'No release notes available'),
                        'source': 'github'
                    }
        
        return {'available': False}
    
    except ImportError:
        print("requests library not installed, skipping GitHub update check")
        return None
    except Exception as e:
        print(f"GitHub update check failed: {e}")
        return None


def check_for_updates_web() -> Optional[Dict]:
    """
    Check for updates on web server
    Requires: pip install requests
    """
    try:
        import requests
        
        response = requests.get(WEB_SERVER_URL, timeout=5)
        
        if response.status_code == 200:
            data = response.json()
            latest_version = data['version']
            
            if ver.parse(latest_version) > ver.parse(__version__):
                return {
                    'available': True,
                    'version': latest_version,
                    'download_url': data['download_url'],
                    'release_notes': data.get('release_notes', 'No release notes available'),
                    'source': 'web'
                }
        
        return {'available': False}
    
    except ImportError:
        print("requests library not installed, skipping web update check")
        return None
    except Exception as e:
        print(f"Web update check failed: {e}")
        return None


def check_for_updates() -> Optional[Dict]:
    """
    Check for updates from all configured sources
    Returns info about the latest available update
    """
    if not UPDATE_CHECK_ENABLED:
        return {'available': False}
    
    # Try network share first (fastest for internal use)
    update_info = check_for_updates_network_share()
    if update_info and update_info.get('available'):
        return update_info
    
    # Try GitHub
    update_info = check_for_updates_github()
    if update_info and update_info.get('available'):
        return update_info
    
    # Try web server
    update_info = check_for_updates_web()
    if update_info and update_info.get('available'):
        return update_info
    
    return {'available': False}


def get_current_version() -> str:
    """Get the current version of the application"""
    return __version__


if __name__ == "__main__":
    # Test the updater
    print(f"Current version: {__version__}")
    print("Checking for updates...")
    
    update_info = check_for_updates()
    
    if update_info.get('available'):
        print(f"\n✓ Update available!")
        print(f"  New version: {update_info['version']}")
        print(f"  Source: {update_info['source']}")
        print(f"\n  Release notes:")
        print(f"  {update_info['release_notes']}")
        
        if 'download_path' in update_info:
            print(f"\n  Download from: {update_info['download_path']}")
        elif 'download_url' in update_info:
            print(f"\n  Download from: {update_info['download_url']}")
    else:
        print("\n✓ You have the latest version!")
