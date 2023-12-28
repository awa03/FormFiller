#!/bin/bash

# Check if Homebrew is installed, and install it if not
if ! command -v brew &> /dev/null; then
    echo "Homebrew not found. Installing Homebrew..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
fi

# Install Python and pip using Homebrew
brew install python

# Install required Python packages
pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client
