#!/bin/bash
# CI/CD script for running tests and generating coverage reports
# Usage: ./scripts/run_tests.sh

set -e

echo "================================"
echo "Lahimena Tours - Test Suite"
echo "================================"
echo ""

# Check if pytest is installed
if ! command -v pytest &> /dev/null; then
    echo "Installing test dependencies..."
    pip install -r requirements-test.txt
fi

echo "Running unit tests..."
pytest tests/ -v

echo ""
echo "Generating coverage report..."
pytest --cov=utils --cov=models --cov-report=term-missing tests/

echo ""
echo "================================"
echo "Tests completed successfully!"
echo "================================"
