the process to get the contacts from the mail id

Go to : Open Google Cloud Console
https://console.cloud.google.com/

search for : Google People API

enable the api
create the credientials
in the step 3 OAuth Client ID? use the desktop app
then save all-> download the credientials.json

go to API and Services-> OAuth consent screen -> verification center
        in it click on " view auidence configuration " 
        click on the add user 0f Test users
        enter your mail id and click submit


next the python process
 folder strucure:
    your_project_folder/
│
├── credentials.json
└── fetch_contacts.py   (we will create this next)
|__Complete_Google_Contacts.xlsx

in the credientials.json-> paste the entire json content which we have downloaded before

pip install -r requirements.txt

run the program
    you will naviate to the access center of the google mail id -> give the access
    
the entire contacts in your mail id will saved in the "Complete_Google_Contacts.xlsx" 