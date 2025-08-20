"""
Email configuration module for customizable email styling.
Loads styling configuration from YAML file or provides defaults.
"""

import yaml
from pathlib import Path
from typing import Dict, Any, Optional


def load_email_config() -> Dict[str, Any]:
    """
    Load email styling configuration from config.yaml file.
    
    Searches for config files in the following order:
    1. /app/config/config.yaml (Docker production path)
    2. ./config/config.yaml (local development)
    3. ./config.yaml (project root)
    
    Returns:
        Dict containing email configuration with all necessary styling options.
        If no config file is found, returns default configuration.
    """
    
    # Default configuration - matches current hardcoded styles
    default_config = {
        "email": {
            "styles": {
                "base": {
                    "body": {
                        "font-family": "Arial, sans-serif",
                        "font-size": "10pt",
                        "color": "rgb(0, 20, 137)",
                        "line-height": "1.4"
                    }
                },
                "elements": {
                    "h2": {
                        "font-weight": "bold",
                        "font-size": "10pt", 
                        "margin": "8px 0"
                    },
                    "h3": {
                        "text-decoration": "underline",
                        "font-size": "10pt",
                        "font-weight": "normal",
                        "margin": "8px 0"
                    },
                    "p": {
                        "margin": "8px 0"
                    },
                    "ul": {
                        "margin": "8px 0",
                        "padding-left": "20px"
                    },
                    "ol": {
                        "margin": "8px 0", 
                        "padding-left": "20px"
                    },
                    "li": {
                        "margin": "3px 0"
                    }
                }
            }
        }
    }
    
    # Potential config file locations
    config_paths = [
        Path("/app/config/config.yaml"),
        Path("./config/config.yaml"),
        Path("./config.yaml")
    ]
    
    # Try to load config from each location
    for config_path in config_paths:
        try:
            if config_path.exists():
                print(f"Loading email config from: {config_path}")
                with open(config_path, 'r', encoding='utf-8') as f:
                    loaded_config = yaml.safe_load(f)
                    
                # Merge loaded config with defaults (loaded values override defaults)
                if loaded_config and 'email' in loaded_config:
                    merged_config = _merge_configs(default_config, loaded_config)
                    return merged_config
                    
        except Exception as e:
            print(f"Warning: Could not load config from {config_path}: {e}")
            continue
    
    # No config file found or loaded, return defaults
    print("Using default email styling configuration")
    return default_config


def _merge_configs(default: Dict[str, Any], loaded: Dict[str, Any]) -> Dict[str, Any]:
    """
    Recursively merge loaded config with defaults.
    Loaded config values override defaults, but missing keys use defaults.
    """
    result = default.copy()
    
    for key, value in loaded.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = _merge_configs(result[key], value)
        else:
            result[key] = value
    
    return result


def generate_email_css(config: Dict[str, Any]) -> str:
    """
    Generate CSS string from email configuration.
    
    Args:
        config: Email configuration dictionary
        
    Returns:
        CSS string ready to be embedded in email HTML
    """
    styles = config.get("email", {}).get("styles", {})
    base_styles = styles.get("base", {})
    element_styles = styles.get("elements", {})
    
    css_lines = []
    
    # Generate base styles (body)
    for selector, properties in base_styles.items():
        css_lines.append(_generate_css_rule(selector, properties))
    
    # Generate element styles
    for selector, properties in element_styles.items():
        css_lines.append(_generate_css_rule(selector, properties))
    
    return "\n".join(css_lines)


def _generate_css_rule(selector: str, properties: Dict[str, str]) -> str:
    """
    Generate a single CSS rule from selector and properties.
    
    Args:
        selector: CSS selector (e.g., 'body', 'h2')
        properties: Dictionary of CSS properties and values
        
    Returns:
        Formatted CSS rule string
    """
    if not properties:
        return ""
    
    prop_lines = []
    for prop, value in properties.items():
        prop_lines.append(f"    {prop}: {value};")
    
    return f"{selector} {{\n" + "\n".join(prop_lines) + "\n}"


def get_email_config() -> Dict[str, Any]:
    """
    Get the current email configuration.
    This is the main function to be used by other modules.
    """
    return load_email_config()