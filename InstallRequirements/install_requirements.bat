@echo off
set ENV_NAME=excel_test_env

echo Creating conda environment %ENV_NAME%...
call conda create -y -n %ENV_NAME% python=3.10 pandas openpyxl pytest pytest-cov

call conda activate %ENV_NAME%
echo Environment %ENV_NAME% is ready. You can now run tests with:
echo python run_tests.py
