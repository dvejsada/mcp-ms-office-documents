"""Utility functions for handling images in PowerPoint presentations.

This module provides functionality to download and validate images from URLs
for embedding in PowerPoint slides.

Security features:
- SSRF protection: blocks internal IP ranges and cloud metadata endpoints
- Content-type validation: only allows image MIME types
- Size limits: prevents resource exhaustion

Configuration:
- SSRF protection settings are loaded from config/ssrf_protection.yaml
- Environment variables can override/extend the configuration:
  - SSRF_PRESET: Select a preset (strict, cloud, development, disabled)
  - SSRF_ADDITIONAL_BLOCKED_HOSTS: Comma-separated additional hostnames
  - SSRF_ADDITIONAL_BLOCKED_NETWORKS: Comma-separated additional CIDR networks
  - SSRF_DISABLE_PROTECTION: Set to 'true' to disable (NOT RECOMMENDED)
"""

import io
import ipaddress
import logging
import os
import socket
from pathlib import Path
from typing import Tuple, Set, List, Optional
from urllib.parse import urlparse

import requests

try:
    import yaml
    YAML_AVAILABLE = True
except ImportError:
    YAML_AVAILABLE = False

logger = logging.getLogger(__name__)

# Allowed image MIME types
ALLOWED_MIME_TYPES = {
    'image/png',
    'image/jpeg',
    'image/jpg',
    'image/gif',
    'image/bmp',
    'image/webp',
    'image/tiff',
}

# Maximum image size in bytes (10 MB)
MAX_IMAGE_SIZE = 10 * 1024 * 1024

# Request timeout in seconds
REQUEST_TIMEOUT = 30


class SSRFConfig:
    """SSRF Protection configuration loaded from YAML or defaults."""
    
    DEFAULT_BLOCKED_HOSTNAMES = {
        'localhost', 'localhost.localdomain',
        'metadata.google.internal', 'metadata.internal',
        'kubernetes.default.svc', 'kubernetes.default', 'kubernetes',
        'docker', 'docker.internal', 'host.docker.internal', 'gateway.docker.internal',
    }
    
    DEFAULT_BLOCKED_SUFFIXES = {'.localhost', '.local', '.internal', '.svc', '.pod'}
    
    DEFAULT_BLOCKED_NETWORKS = [
        '127.0.0.0/8', '::1/128',
        '10.0.0.0/8', '172.16.0.0/12', '192.168.0.0/16',
        '169.254.0.0/16', 'fe80::/10',
        '100.64.0.0/10', '0.0.0.0/8', '::/128',
        '192.0.2.0/24', '198.51.100.0/24', '203.0.113.0/24', 'fc00::/7',
    ]
    
    def __init__(self):
        self.blocked_hostnames: Set[str] = set()
        self.blocked_suffixes: Set[str] = set()
        self.blocked_networks: List = []
        self.protection_enabled: bool = True
        self.preset_name: str = 'strict'
        self._load_config()
    
    def _find_config_file(self) -> Optional[Path]:
        for path in [
            Path('/app/config/ssrf_protection.yaml'),  # User override in Docker
            Path('config/ssrf_protection.yaml'),  # User override in local dev
            Path('/app/default_config/ssrf_protection.yaml'),  # Default in Docker
            Path(__file__).parent.parent / 'default_config' / 'ssrf_protection.yaml',  # Default relative
            Path('default_config/ssrf_protection.yaml'),  # Default in local dev
        ]:
            if path.exists():
                return path
        return None
    
    def _load_config(self):
        if os.environ.get('SSRF_DISABLE_PROTECTION', '').lower() == 'true':
            logger.warning("[SECURITY] SSRF protection is DISABLED!")
            self.protection_enabled = False
            return
        
        config_file = self._find_config_file()
        config_data = None
        
        if config_file and YAML_AVAILABLE:
            try:
                with open(config_file, 'r') as f:
                    config_data = yaml.safe_load(f)
                logger.info(f"[SSRF] Loaded config from {config_file}")
            except Exception as e:
                logger.warning(f"[SSRF] Failed to load config: {e}")
        
        preset_name = os.environ.get('SSRF_PRESET', '')
        if not preset_name and config_data:
            preset_name = config_data.get('default_preset', 'strict')
        self.preset_name = preset_name or 'strict'
        
        enabled_categories = None
        if config_data and 'presets' in config_data:
            preset = config_data['presets'].get(self.preset_name, {})
            enabled_categories = set(preset.get('enabled_categories', []))
            logger.info(f"[SSRF] Using preset '{self.preset_name}'")
        
        if config_data and 'blocked_hostnames' in config_data:
            self._load_hostnames(config_data['blocked_hostnames'], enabled_categories)
        else:
            self.blocked_hostnames = self.DEFAULT_BLOCKED_HOSTNAMES.copy()
            self.blocked_suffixes = self.DEFAULT_BLOCKED_SUFFIXES.copy()
        
        if config_data and 'blocked_networks' in config_data:
            self._load_networks(config_data['blocked_networks'], enabled_categories)
        else:
            for cidr in self.DEFAULT_BLOCKED_NETWORKS:
                try:
                    self.blocked_networks.append(ipaddress.ip_network(cidr))
                except ValueError:
                    pass
        
        self._add_env_overrides()
        logger.info(f"[SSRF] {len(self.blocked_hostnames)} hosts, {len(self.blocked_networks)} networks")
    
    def _load_hostnames(self, config, enabled):
        for cat_name, cat_data in config.items():
            if enabled and cat_name not in enabled:
                continue
            for entry in cat_data.get('entries', []):
                if entry.startswith('.'):
                    self.blocked_suffixes.add(entry.lower())
                else:
                    self.blocked_hostnames.add(entry.lower())
    
    def _load_networks(self, config, enabled):
        for cat_name, cat_data in config.items():
            if enabled and cat_name not in enabled:
                continue
            for cidr in cat_data.get('ipv4', []) + cat_data.get('ipv6', []):
                try:
                    self.blocked_networks.append(ipaddress.ip_network(cidr))
                except ValueError as e:
                    logger.warning(f"[SSRF] Invalid network '{cidr}': {e}")
    
    def _add_env_overrides(self):
        for host in os.environ.get('SSRF_ADDITIONAL_BLOCKED_HOSTS', '').split(','):
            host = host.strip().lower()
            if host:
                (self.blocked_suffixes if host.startswith('.') else self.blocked_hostnames).add(host)
        for cidr in os.environ.get('SSRF_ADDITIONAL_BLOCKED_NETWORKS', '').split(','):
            if cidr.strip():
                try:
                    self.blocked_networks.append(ipaddress.ip_network(cidr.strip()))
                except ValueError:
                    pass


_ssrf_config = SSRFConfig()


class ImageDownloadError(Exception):
    pass

class ImageValidationError(Exception):
    pass

class SSRFProtectionError(Exception):
    pass


def _is_ip_blocked(ip_str: str) -> bool:
    if not _ssrf_config.protection_enabled:
        return False
    try:
        ip = ipaddress.ip_address(ip_str)
        return any(ip in net for net in _ssrf_config.blocked_networks)
    except ValueError:
        return False


def _is_hostname_blocked(hostname: str) -> bool:
    if not _ssrf_config.protection_enabled:
        return False
    hostname_lower = hostname.lower()
    if hostname_lower in _ssrf_config.blocked_hostnames:
        return True
    return any(hostname_lower.endswith(s) for s in _ssrf_config.blocked_suffixes)


def _resolve_hostname(hostname: str) -> list:
    try:
        addr_info = socket.getaddrinfo(hostname, None, socket.AF_UNSPEC)
        return list(set(info[4][0] for info in addr_info))
    except socket.gaierror as e:
        raise SSRFProtectionError(f"Cannot resolve '{hostname}': {e}")


def is_ssrf_safe(url: str) -> Tuple[bool, str]:
    if not _ssrf_config.protection_enabled:
        return True, ""
    try:
        parsed = urlparse(url)
        hostname = parsed.hostname
        if not hostname:
            return False, "URL has no hostname"
        if _is_hostname_blocked(hostname):
            return False, f"Blocked hostname: {hostname}"
        try:
            if _is_ip_blocked(hostname):
                return False, f"Blocked IP: {hostname}"
            return True, ""
        except ValueError:
            pass
        try:
            for ip_str in _resolve_hostname(hostname):
                if _is_ip_blocked(ip_str):
                    return False, f"'{hostname}' resolves to blocked IP: {ip_str}"
        except SSRFProtectionError as e:
            return False, str(e)
        return True, ""
    except Exception as e:
        return False, f"Error: {e}"


def validate_url(url: str) -> bool:
    try:
        parsed = urlparse(url)
        if parsed.scheme not in ('http', 'https') or not parsed.netloc:
            return False
        is_safe, error = is_ssrf_safe(url)
        if not is_safe:
            logger.warning(f"SSRF blocked: {url} - {error}")
            return False
        return True
    except Exception:
        return False


def get_ssrf_config() -> SSRFConfig:
    return _ssrf_config


def reload_ssrf_config():
    global _ssrf_config
    _ssrf_config = SSRFConfig()


def download_image(url: str) -> Tuple[io.BytesIO, str]:
    try:
        parsed = urlparse(url)
        if parsed.scheme not in ('http', 'https') or not parsed.netloc:
            raise ImageValidationError(f"Invalid URL: {url}")
    except Exception:
        raise ImageValidationError(f"Invalid URL: {url}")
    
    is_safe, error = is_ssrf_safe(url)
    if not is_safe:
        logger.warning(f"SSRF blocked download: {url} - {error}")
        raise SSRFProtectionError(f"Blocked: {error}")
    
    logger.info(f"Downloading image: {url}")
    
    try:
        response = requests.get(url, timeout=REQUEST_TIMEOUT, stream=True,
            headers={'User-Agent': 'Mozilla/5.0 PowerPoint-MCP/1.0'},
            allow_redirects=False)
        
        redirect_count = 0
        while response.is_redirect and redirect_count < 5:
            redirect_url = response.headers.get('Location')
            if not redirect_url:
                break
            if not redirect_url.startswith(('http://', 'https://')):
                from urllib.parse import urljoin
                redirect_url = urljoin(url, redirect_url)
            is_safe, error = is_ssrf_safe(redirect_url)
            if not is_safe:
                raise SSRFProtectionError(f"Redirect blocked: {redirect_url} - {error}")
            response = requests.get(redirect_url, timeout=REQUEST_TIMEOUT, stream=True,
                headers={'User-Agent': 'Mozilla/5.0 PowerPoint-MCP/1.0'},
                allow_redirects=False)
            redirect_count += 1
            url = redirect_url
        response.raise_for_status()
    except SSRFProtectionError:
        raise
    except requests.exceptions.Timeout:
        raise ImageDownloadError(f"Timeout: {url}")
    except requests.exceptions.ConnectionError:
        raise ImageDownloadError(f"Connection error: {url}")
    except requests.exceptions.HTTPError as e:
        raise ImageDownloadError(f"HTTP {e.response.status_code}: {url}")
    except requests.exceptions.RequestException as e:
        raise ImageDownloadError(f"Error: {url}: {e}")
    
    content_type = response.headers.get('Content-Type', '').split(';')[0].strip().lower()
    if content_type and content_type not in ALLOWED_MIME_TYPES:
        raise ImageValidationError(f"Invalid type: {content_type}")
    
    content_length = response.headers.get('Content-Length')
    if content_length:
        try:
            if int(content_length) > MAX_IMAGE_SIZE:
                raise ImageValidationError(f"Too large: {int(content_length)/(1024*1024):.1f}MB")
        except ValueError:
            pass
    
    image_data = io.BytesIO()
    total_size = 0
    for chunk in response.iter_content(chunk_size=8192):
        total_size += len(chunk)
        if total_size > MAX_IMAGE_SIZE:
            raise ImageValidationError(f"Too large (>{MAX_IMAGE_SIZE/(1024*1024):.0f}MB)")
        image_data.write(chunk)
    image_data.seek(0)
    
    extension = get_image_extension(content_type, url)
    logger.info(f"Downloaded: {total_size/1024:.1f}KB, type: {extension}")
    return image_data, extension


def get_image_extension(content_type: str, url: str) -> str:
    type_to_ext = {
        'image/png': 'png', 'image/jpeg': 'jpg', 'image/jpg': 'jpg',
        'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'webp', 'image/tiff': 'tiff',
    }
    if content_type in type_to_ext:
        return type_to_ext[content_type]
    path = urlparse(url).path.lower()
    for ext in ('png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'tiff'):
        if path.endswith(f'.{ext}'):
            return 'jpg' if ext == 'jpeg' else ext
    return 'png'
