"""
Validation utilities for the application
"""

try:
    from phonenumbers import parse, is_valid_number
    PHONENUMBERS_AVAILABLE = True
except ImportError:
    PHONENUMBERS_AVAILABLE = False

import re


def get_exchange_rates():
    """
    Get current exchange rates for EUR and USD to MGA

    Returns:
        dict: Dictionary with rates {'EUR': rate, 'USD': rate}
        where rate is how many MGA for 1 EUR/USD
    """
    try:
        import requests
        # Using exchangerate-api.com free API with MGA as base
        response = requests.get('https://api.exchangerate-api.com/v4/latest/MGA', timeout=5)
        data = response.json()
        # Convert to more readable format: how many MGA for 1 EUR/USD
        rates = {
            'EUR': 1 / data['rates']['EUR'],  # 1 EUR = X MGA
            'USD': 1 / data['rates']['USD']   # 1 USD = X MGA
        }
        return rates
    except Exception as e:
        # Fallback to hardcoded rates if API fails
        print(f"Warning: Could not fetch exchange rates: {e}. Using fallback rates.")
        return {
            'EUR': 5235.0,  # Approximate 1 EUR = 5235 MGA
            'USD': 4900.0   # Approximate 1 USD = 4900 MGA
        }


def convert_currency(amount, from_currency, to_currency, rates=None):
    """
    Convert amount from one currency to another

    Args:
        amount (float): Amount to convert
        from_currency (str): Source currency ('Ariary', 'Euro', 'Dollar US')
        to_currency (str): Target currency ('Ariary', 'Euro', 'Dollar US')
        rates (dict): Exchange rates dict, if None will fetch

    Returns:
        float: Converted amount
    """
    if from_currency == to_currency:
        return amount

    if rates is None:
        rates = get_exchange_rates()

    # Convert to MGA first if needed
    if from_currency == 'Euro':
        amount_mga = amount * rates['EUR']  # EUR to MGA
    elif from_currency == 'Dollar US':
        amount_mga = amount * rates['USD']  # USD to MGA
    else:  # Ariary
        amount_mga = amount

    # Convert to target currency
    if to_currency == 'Euro':
        return amount_mga / rates['EUR']  # MGA to EUR
    elif to_currency == 'Dollar US':
        return amount_mga / rates['USD']  # MGA to USD
    else:  # Ariary
        return amount_mga


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