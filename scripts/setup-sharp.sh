#!/bin/bash

# Sharp prebuild script for GitHub Actions
echo "Setting up Sharp prebuilt binaries..."

# Set environment variables for Sharp
export npm_config_sharp_binary_host="https://github.com/lovell/sharp-libvips/releases/download/"
export npm_config_sharp_libvips_binary_host="https://github.com/lovell/sharp-libvips/releases/download/"

# Force Sharp to use prebuilt binaries
export npm_config_sharp_ignore_global_libvips="true"

# Install Sharp with specific platform binaries
case "$(uname -s)" in
    MINGW*|CYGWIN*|MSYS*)
        echo "Windows detected - using win32-x64 Sharp binary"
        npm install --platform=win32 --arch=x64 sharp
        ;;
    Darwin*)
        echo "macOS detected - using darwin-x64 Sharp binary"
        npm install --platform=darwin --arch=x64 sharp
        ;;
    Linux*)
        echo "Linux detected - using linux-x64 Sharp binary"
        npm install --platform=linux --arch=x64 sharp
        ;;
    *)
        echo "Unknown platform, using default Sharp installation"
        npm install sharp
        ;;
esac

echo "Sharp setup completed"
