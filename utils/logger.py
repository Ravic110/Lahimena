"""
Logging configuration for Lahimena Tours application
"""

import logging
import os
from datetime import datetime


def setup_logger():
    """
    Setup application-wide logger with file and console handlers
    
    Returns:
        logging.Logger: Configured logger instance
    """
    # Create logs directory if it doesn't exist
    log_dir = os.path.join(os.path.dirname(__file__), '..', 'logs')
    os.makedirs(log_dir, exist_ok=True)
    
    # Generate log filename with date
    log_filename = os.path.join(
        log_dir,
        f"app_{datetime.now().strftime('%Y%m%d')}.log"
    )
    
    # Create logger
    logger = logging.getLogger('lahimena_tours')
    
    # Set logging level
    logger.setLevel(logging.DEBUG)
    
    # Remove existing handlers to avoid duplicates
    logger.handlers.clear()
    
    # Create formatters
    file_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    console_formatter = logging.Formatter(
        '%(levelname)-8s - %(message)s'
    )
    
    # Create file handler
    file_handler = logging.FileHandler(log_filename)
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(file_formatter)
    
    # Create console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(console_formatter)
    
    # Add handlers to logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger


# Create global logger instance
logger = setup_logger()


# Convenience functions for common logging tasks
def log_info(message):
    """Log info message"""
    logger.info(message)


def log_warning(message):
    """Log warning message"""
    logger.warning(message)


def log_error(message, exc_info=False):
    """Log error message"""
    logger.error(message, exc_info=exc_info)


def log_critical(message, exc_info=False):
    """Log critical message"""
    logger.critical(message, exc_info=exc_info)


def log_debug(message):
    """Log debug message"""
    logger.debug(message)
