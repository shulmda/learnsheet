rm -R dist
pyinstaller --hidden-import wx --onefile --windowed ./learnsheet.py
rm -f learnsheet.zip 
mkdir learnsheet
cp -R dist/ learnsheet
cp ./Info.plist dist/learnsheet.app/Contents/
cp learnsheet.icns dist/learnsheet.app/Contents/Resources/icon-windowed.icns
zip learnsheet.zip learnsheet
rm -rf learnsheet
