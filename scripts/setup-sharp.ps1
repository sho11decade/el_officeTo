# Sharp setup script for Windows PowerShell
Write-Host "Setting up Sharp prebuilt binaries for Windows..."

# Set environment variables for Sharp
$env:npm_config_sharp_binary_host = "https://github.com/lovell/sharp-libvips/releases/download/"
$env:npm_config_sharp_libvips_binary_host = "https://github.com/lovell/sharp-libvips/releases/download/"
$env:npm_config_sharp_ignore_global_libvips = "true"

# Force Sharp to use prebuilt binaries for Windows x64
Write-Host "Installing Sharp with win32-x64 prebuilt binary..."
npm install --platform=win32 --arch=x64 sharp

Write-Host "Sharp setup completed successfully"
