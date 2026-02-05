"""
Validation utilities for the application
"""

try:
    from phonenumbers import parse, is_valid_number
    PHONENUMBERS_AVAILABLE = True
except ImportError:
    PHONENUMBERS_AVAILABLE = False

import re
from utils.logger import logger


def validate_client_reference(ref_client):
    """
    Validate client reference format
    
    Args:
        ref_client (str): Client reference to validate
        
    Returns:
        tuple: (is_valid: bool, error_message: str or None)
    """
    if not ref_client or not ref_client.strip():
        return False, "La référence client est requise"
    
    # Allow alphanumeric, hyphens, and underscores
    if not re.match(r'^[A-Za-z0-9_-]{3,20}$', ref_client.strip()):
        return False, "Référence: 3-20 caractères (lettres, chiffres, - et _)"
    
    return True, None


def validate_hotel_name(nom):
    """
    Validate hotel name
    
    Args:
        nom (str): Hotel name to validate
        
    Returns:
        tuple: (is_valid: bool, error_message: str or None)
    """
    if not nom or not nom.strip():
        return False, "Le nom de l'hôtel est requis"
    
    if len(nom.strip()) < 2:
        return False, "Le nom doit avoir au moins 2 caractères"
    
    if len(nom.strip()) > 100:
        return False, "Le nom ne peut pas dépasser 100 caractères"
    
    return True, None


def validate_required_field(value, field_name):
    """
    Validate that a field is not empty
    
    Args:
        value (str): Value to validate
        field_name (str): Name of field for error message
        
    Returns:
        tuple: (is_valid: bool, error_message: str or None)
    """
    if not value or not str(value).strip():
        return False, f"{field_name} est requis"
    
    return True, None


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
    Validate email address format with comprehensive checks
    
    Args:
        email (str): Email address to validate
        
    Returns:
        tuple: (is_valid: bool, error_message: str or None)
    """
    if not email:
        return False, "L'email est requis"
    
    email = email.strip()
    
    # Check email length
    if len(email) > 254:
        return False, "Email trop long (max 254 caractères)"
    
    # Check email format
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    if not re.match(pattern, email):
        return False, "Format email invalide"
    
    # Check for consecutive dots
    if '..' in email:
        return False, "Email invalide (points consécutifs)"
    
    # Check local part length (before @)
    local_part = email.split('@')[0]
    if len(local_part) > 64:
        return False, "Partie locale de l'email trop longue"
    
    logger.debug(f"Email validated: {email}")
    return True, None


def validate_phone_number(phone_code, phone_number):
    """
    Validate phone number using phonenumbers library with comprehensive checks
    
    Args:
        phone_code (str): Country code (e.g., "+261" or "MG")
        phone_number (str): Phone number to validate
        
    Returns:
        tuple: (is_valid: bool, error_message: str or None)
    """
    if not phone_number or not str(phone_number).strip():
        return False, "Le numéro de téléphone est requis"
    
    phone_number = str(phone_number).strip()
    
    # Remove common separators
    phone_number = phone_number.replace(" ", "").replace("-", "").replace(".", "")
    
    # Check minimum length
    if len(phone_number) < 7:
        return False, "Numéro de téléphone trop court"
    
    # Check maximum length
    if len(phone_number) > 20:
        return False, "Numéro de téléphone trop long"
    
    if not PHONENUMBERS_AVAILABLE:
        # Basic validation if phonenumbers is not available
        logger.warning("phonenumbers library not available, using basic validation")
        return True, None
    
    try:
        # Normalize phone code if needed
        if phone_code and not phone_code.startswith('+'):
            # Assume it's country code like 'MG', convert to '+261'
            country_map = {
                'MG': '+261',
                'FR': '+33',
                'US': '+1'
            }
            full_code = country_map.get(phone_code.upper(), '+' + phone_code)
        else:
            full_code = phone_code or '+261'  # Default to Madagascar
        
        full_number = full_code + phone_number
        parsed = parse(full_number)
        
        if is_valid_number(parsed):
            logger.debug(f"Phone number validated: {phone_code}:{phone_number}")
            return True, None
        else:
            return False, "Numéro de téléphone invalide pour ce pays"
    except Exception as e:
        logger.warning(f"Phone validation error: {e}")
        return False, f"Erreur validation téléphone: {str(e)}"


def validate_price(price_str):
    """
    Validate and parse price value
    
    Args:
        price_str (str): Price as string
        
    Returns:
        tuple: (is_valid: bool, price_value: float or None, error_message: str or None)
    """
    if not price_str or not str(price_str).strip():
        return False, None, "Le prix est requis"
    
    price_str = str(price_str).strip()
    
    try:
        # Replace common separators
        price_str = price_str.replace(',', '.').replace(' ', '')
        price = float(price_str)
        
        if price < 0:
            return False, None, "Le prix ne peut pas être négatif"
        
        if price > 999999999:
            return False, None, "Le prix est trop élevé"
        
        logger.debug(f"Price validated: {price}")
        return True, price, None
    except ValueError:
        return False, None, "Format de prix invalide"


def validate_date_format(date_str):
    """
    Validate date format (DD/MM/YYYY)
    
    Args:
        date_str (str): Date string to validate
        
    Returns:
        tuple: (is_valid: bool, error_message: str or None)
    """
    if not date_str or not str(date_str).strip():
        return False, "La date est requise"
    
    date_str = str(date_str).strip()
    
    # Check format
    pattern = r'^\d{1,2}/\d{1,2}/\d{4}$'
    if not re.match(pattern, date_str):
        return False, "Format de date invalide (DD/MM/YYYY)"
    
    # Validate date ranges
    try:
        parts = date_str.split('/')
        day, month, year = int(parts[0]), int(parts[1]), int(parts[2])
        
        if month < 1 or month > 12:
            return False, "Mois invalide (1-12)"
        
        if day < 1 or day > 31:
            return False, "Jour invalide (1-31)"
        
        if year < 2000 or year > 2100:
            return False, "Année invalide (2000-2100)"
        
        logger.debug(f"Date validated: {date_str}")
        return True, None
    except Exception as e:
        return False, f"Erreur validation date: {str(e)}"


def validate_integer(value_str, min_val=None, max_val=None):
    """
    Validate and parse integer value
    
    Args:
        value_str (str): Integer as string
        min_val (int): Minimum value (optional)
        max_val (int): Maximum value (optional)
        
    Returns:
        tuple: (is_valid: bool, value: int or None, error_message: str or None)
    """
    if not value_str or not str(value_str).strip():
        return False, None, "La valeur est requise"
    
    try:
        value = int(str(value_str).strip())
        
        if min_val is not None and value < min_val:
            return False, None, f"Valeur trop faible (min: {min_val})"
        
        if max_val is not None and value > max_val:
            return False, None, f"Valeur trop élevée (max: {max_val})"
        
        return True, value, None
    except ValueError:
        return False, None, "Format entier invalide"