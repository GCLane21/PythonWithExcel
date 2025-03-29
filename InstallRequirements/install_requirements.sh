#!/bin/bash

# Check if conda is installed
if ! command -v conda &> /dev/null; then
    echo "Conda is not installed. Please install Anaconda or Miniconda."
    exit 1
fi

# Define the environment name
ENV_NAME="excel_test_env"

# Create a new conda environment with required packages
echo "Creating conda environment '$ENV_NAME'..."
conda create -y -n $ENV_NAME python=3.10 pandas openpyxl pytest pytest-cov

# Activate the environment
source $(conda info --base)/etc/profile.d/conda.sh
conda activate $ENV_NAME

# Confirmation message
echo "Environment '$ENV_NAME' is ready. You can now run tests with coverage using:"
echo "python run_tests.py"
