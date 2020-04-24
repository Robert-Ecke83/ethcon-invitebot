from flask import Flask, request
from webexteamssdk import WebexTeamsAPI
from webexteamssdk import exceptions
import argparse
import json
import requests
import csv

app = Flask(__name__)


token = 'NDE1Y2E5OGYtMzQ3MC00NDEwLWE3YjQtYzNlMDdjZmIyMWM2OGJlYWE5OTAtYzg4_PE93_e9ed1911-dd0d-4e14-a181-b40037bebc6b'

api = WebexTeamsAPI(access_token=token)


@app.route('/')
def hello():
    return 'You made it, I´m still running!'

# Receive POST from Spark Space
@app.route('/sparkhook', methods=['POST'])
def sparkhook():

    if request.method == 'POST':

        jsonAnswer = json.loads(request.data) # Format data from POST into JSON

        botDetails = api.people.me() # Get details of the bot from its token

        if str(jsonAnswer['data']['personEmail']) != str(botDetails.emails[0]): # If the message is not sent by the bot

            botName = str(botDetails.displayName) # Get bot's display name
            botFirstName = botName.split(None, 1)[0] # Get bot's "first name"

            sparkMessage = api.messages.get(jsonAnswer['data']['id']) # Get message object text from message ID
            sparkMsgText = str(sparkMessage.text) # Get message text
            sparkMsgRoomId = str(sparkMessage.roomId) # Get message roomId
            sparkMsgPersonEmail = str(sparkMessage.personEmail) # Get message personEmail

            if "hello" in sparkMsgText: #Replies to the message hello
                textAnswer = 'Hello <@personEmail:' + str(jsonAnswer['data']['personEmail']) + '> please @mention me with help'
                botAnswered = api.messages.create(roomId=sparkMsgRoomId, markdown=textAnswer)

            elif "help" in sparkMsgText: #Replies to the message help
                textAnswer = 'Hello <@personEmail:' + str(jsonAnswer['data']['personEmail']) + '>, I will help you to invite People to a Webex Teams Space or Team: \n - In order to invite people into a space prepare a *.csv and send to me via @mention into the space you want the people to be added \n - If you want me to add people to a Team I need to be **Moderator** in that specific general Team Space \n - Attached you find a example *.csv make sure people in the *.csv are **not** already in the space\team \n - If you want to know more about me @mention me with **about**'
                botAnswered = api.messages.create(roomId=sparkMsgRoomId, markdown=textAnswer, files=["https://raw.githubusercontent.com/Robert-Ecke83/ethcon-invitebot/master/invite.csv"])
            
            elif "about" in sparkMsgText: #Replies to the message about
                textAnswer = 'Hello <@personEmail:' + str(jsonAnswer['data']['personEmail']) + '>, \n - Here you can find the code: https://github.com/Robert-Ecke83/ethcon-invitebot \n - recke@ethcon.de'
                botAnswered = api.messages.create(roomId=sparkMsgRoomId, markdown=textAnswer)
            
            else:

                 if "@" in sparkMsgPersonEmail: # Check if the Message comes from @meetingzone.com user

                    if not sparkMessage.files: #IF there is no CSV file attached use help
                        textAnswer = 'Hello <@personEmail:' + str(jsonAnswer['data']['personEmail']) + '> I´m missing the **CSV** file please @mention me with **help** to find out how to use me'
                        botAnswered = api.messages.create(roomId=sparkMsgRoomId, markdown=textAnswer)

                        # If the message comes with a file
                    else:
                        sparkMsgFileUrl = str(sparkMessage.files[0]) # Get the URL of the first file
                        
                        sparkHeader = {'Authorization': "Bearer " + token}
                        i = 0 # Index to skip first row in the CSV file

                        with requests.Session() as s: # Creating a session to allow several HTTP messages with one TCP connection
                            getResponse = s.get(sparkMsgFileUrl, headers=sparkHeader, timeout=30.0) # Get file

                            # If the file extension is CSV
                            if str(getResponse.headers['Content-Type']) == 'text/csv':
                                decodedContent = getResponse.content.decode('utf-8')
                                csvFile = csv.reader(decodedContent.splitlines(), delimiter=';')
                                listEmails = list(csvFile)
                                textAnswer = 'Hello <@personEmail:' + str(jsonAnswer['data']['personEmail']) + '>, I will start the Invite now' #Message to Room Invite will start
                                botAnswered = api.messages.create(roomId=sparkMsgRoomId, markdown=textAnswer)
                                for row in listEmails: # Creating one list for each line in the file
                                    try:
                                        api.memberships.create(roomId=sparkMsgRoomId, personEmail=str(row[2]))
                                        
                                    except exceptions.ApiError as e:
                                        if e.response.status_code == 409:
                                                                                
                            else:   # If the attached file is not a CSV
                                textAnswer = 'Sorry, I only understand **CSV** files, please @mention me with **help** to find out how to use me'
                                botAnswered = api.messages.create(roomId=sparkMsgRoomId, markdown=textAnswer)

                 else:
                    textAnswer = 'Sorry, I´m only allowed to invite people for MeetingZone Employes.'
                    botAnswered = api.messages.create(roomId=sparkMsgRoomId, markdown=textAnswer)

    return 'OK'

if __name__ == '__main__':
    app.run()
