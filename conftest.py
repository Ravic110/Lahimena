"""
Pytest configuration and fixtures for Lahimena Tours tests
"""

import pytest
import sys
import os
from pathlib import Path

# Add the project root to the Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))


@pytest.fixture
def client_data_dict():
    """Fixture providing sample client data"""
    return {
        'ref_client': 'CLI-2026-001',
        'nom': 'Test Client',
        'telephone': '0301234567',
        'email': 'test@example.com',
        'code_pays': 'MG',
        'periode': '5 jours',
        'restauration': 'Petit-déjeuner',
        'hebergement': 'Hôtel 3*',
        'chambre': 'Double',
        'enfant': 'Non',
        'age_enfant': '',
        'forfait': 'Standard',
        'circuit': 'Côte Est'
    }


@pytest.fixture
def hotel_data_dict():
    """Fixture providing sample hotel data"""
    return {
        'nom': 'Test Hotel',
        'adresse': '123 Test Street',
        'email': 'hotel@test.mg',
        'telephone': '0302345678',
        'code_pays': 'MG',
        'etoiles': '3',
        'region': 'Centre'
    }


@pytest.fixture
def temp_excel_file(tmp_path):
    """Fixture providing a temporary Excel file path"""
    return str(tmp_path / "test_data.xlsx")


@pytest.fixture
def mock_logger(monkeypatch):
    """Fixture providing a mock logger"""
    from unittest.mock import MagicMock
    mock = MagicMock()
    
    # Mock the logger functions
    mock.info = MagicMock()
    mock.warning = MagicMock()
    mock.error = MagicMock()
    mock.debug = MagicMock()
    
    return mock


@pytest.fixture(scope="session")
def project_root():
    """Fixture providing the project root path"""
    return Path(__file__).parent.parent


def pytest_configure(config):
    """Configure pytest with custom settings"""
    # Add custom configuration here if needed
    pass
