.PHONY: test test-unit test-cov test-verbose install-test clean

# Run all tests
test:
	pytest

# Run only unit tests
test-unit:
	pytest tests/ -m unit

# Run tests with coverage report
test-cov:
	pytest --cov=utils --cov=models --cov=gui --cov-report=html --cov-report=term

# Run tests in verbose mode
test-verbose:
	pytest -v

# Install test dependencies
install-test:
	pip install -r requirements-test.txt

# Run tests and show coverage
coverage:
	pytest --cov=. --cov-report=html --cov-report=term-missing tests/

# Clean up test artifacts
clean:
	rm -rf .pytest_cache
	rm -rf htmlcov
	rm -rf .coverage
	find . -type d -name __pycache__ -exec rm -rf {} +
	find . -type f -name "*.pyc" -delete

# Run specific test file
test-file:
	@echo "Usage: make test-file FILE=tests/test_validators.py"
	pytest $(FILE) -v

# Run tests matching pattern
test-match:
	@echo "Usage: make test-match PATTERN=test_email"
	pytest -k "$(PATTERN)" -v
