pyinstaller --hidden-import tkinter --onefile --windowed ./learnsheet.py
rm -f learnsheet.zip 
mkdir learnsheet
cp -R dist/ learnsheet
zip learnsheet.zip learnsheet
rm -rf learnsheet
