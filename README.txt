This project is something I have worked on to monitor a gmail account for new emails from Serve/BB to help keep track of card loads. The program is an infinite loop, and is meant to be ran 24/7, although you can start/stop it as you want. Below is setup information. 

You will need to do the following:

Install Python

Then pip install these commands:

pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

pip install spread

pip install beautifulsoup4

pip install oauth2client

Pip install lxml


Then open the file Card_Load_Tracker.py in notepad or your text editor of choice. 

Change line 21 to the path to the folder you have these files stored in. 

Change line 23 to the name of your google sheets sheet. 

Go to https://console.developers.google.com/ 

Create a project. The project can be named anything you want. Once you have the project created, go enable the Gmail API, the google drive API, and the google Sheets API. 

Once you have enabled the 3 API's create Oauth credentials and download the .json file. Rename it to credentials.json

Run Reset_credential.py and allow access for the script. 

You are finished setting up. You should be able to run the Card_Load_tracker.py file successfully now. 