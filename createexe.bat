echo off
cls
echo "Creating..."

pyinstaller --onefile --windowed file_explorer.py

echo "Done!"
pause
