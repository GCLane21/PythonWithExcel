import pytest
import sys

if __name__ == "__main__":
    # Run tests with coverage on ExcelWriter.py and output a coverage report
    sys.exit(pytest.main([
        "--cov=ExcelWriter",
        "--cov-report=term-missing",
        "test_excel_writer.py"
    ]))
