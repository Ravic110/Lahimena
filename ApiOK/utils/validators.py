"""
Validation utilities for the application
"""

try:
    from phonenumbers import parse, is_valid_number
    PHONENUMBERS_AVAILABLE = True
except ImportError:
    PHONENUMBERS_AVAILABLE = False

import re


def validate_email(email):
    """
    Validate email address format

    Args:
        email (str): Email address to validate

    Returns:
        bool: True if valid, False otherwise
    """
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None


def validate_phone_number(phone_code, phone_number):
    """
    Validate phone number using phonenumbers library

    Args:
        phone_code (str): Country code (e.g., "+261")
        phone_number (str): Phone number without country code

    Returns:
        bool: True if valid, False otherwise
    """
    if not PHONENUMBERS_AVAILABLE:
        # Basic validation if phonenumbers is not available
        return len(phone_number.strip()) >= 7

    try:
        full_number = phone_code + phone_number
        parsed = parse(full_number)
        return is_valid_number(parsed)
    except:
        return False