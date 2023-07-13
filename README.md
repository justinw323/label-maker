# label-maker
Generate documents for parts ready to be shipped

# Installation
Install anaconda and set up a vitual environment following these instructions: https://www.geeksforgeeks.org/set-up-virtual-environment-for-python-using-anaconda/# (Optional)

1. Click the green "Code" button, click "Download ZIP"
2. Unzip the files and move to your choice of location

From a terminal:
1. Navigate to the folder containing the .spec file (label_maker.spec or label_maker_fast.spec)
2. Install the pyinstaller package with:
   1. 'pip install pyinstaller'
   2. 'pip3 install pyinstaller'
   3. 'conda install pyinstaller' if using the Anaconda prompt
4. Run:
   1. 'pyinstaller label_maker.spec' for the slower app
   2. 'pyinstaller label_maker_fast.spec' for the faster app

# Usage
1. Fill in an input sheet, refer to the example sheet if necessary
2. Make sure that you have the templates downloaded (CoC Template New.docx and Packing List Template New.docx)
3. Run the .exe file
4. When it opens, choose your template files, your input spreadsheet, and the folder where you want the documents to be saved, and click "Generate Documents"
   1. Close the input spreadsheet when you generate the documents to create the spreadsheet for the next batch of the order
6. If there are any issues, the program will let you know
7. If successful, a folder for the batch will be created containing all of the documents
