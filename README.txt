This is a local website that links to a spreadsheet that tracks your time spent on various categories

Necessary python packages:

flask - pip install flask
flask_cors - pip install flask_cors
openpyxl - pip install openpyxl

The rest of the libraries used are standard python libraries

1. Download the excel spreadsheet, app.py, and index.html and place all in the same folder
2. Configure the excel spreadsheet name, the week, and your categories in app.py
3. Reflect those same categories as row headers in your excel spreadsheet
4. Go into gitbash and open up the folder and type python app.py
5. Go into a browser and open http://localhost:5000


Notes: You will have to change the week as each week passes
For this to run, you will need to keep gitbash open

Additionally I made a python program(DeepLearningTracker.py) that can run locally in vscode and is interactable with commands inside the terminal
Similarly you must change the week as each week passes and configure your categories in DeepLearningTracker.py and the spreadsheet

Terminal commands:
"start {category}" to start a timer for that category
"stop" to stop that timer and add that time to the spreadsheet
"exit" to exit the program
