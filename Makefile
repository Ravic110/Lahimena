.PHONY: test test-fast test-gui test-unit test-cov test-verbose install-test clean lint init-data

# Run all tests, interface comprise (~95 s : les ecrans de cotation
# construisent de vrais widgets)
test:
	pytest

# Suite rapide, sans les tests d'interface (~8 s)
test-fast:
	pytest -m "not gui"

# Seuls les tests d'interface
test-gui:
	pytest -m gui

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

# Lint du depot entier. gui/forms et finances ne sont plus exclus de
# setup.cfg : `flake8 .` les couvre, la cible stricte separee n'a plus d'objet.
lint:
	flake8 . --statistics --count
	black --check .
	isort --check-only .

# Amorcage des classeurs de donnees sur un poste neuf
init-data:
	python scripts/init_data.py init

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
