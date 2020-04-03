from flask import Flask, request
from ciscosparkapi import CiscoSparkAPI
import json
import requests
import csv

app = Flask(__name__)


BOT_TOKEN = 'NmNhYjkyNTctZmUyYy00NmI0LTlkM2MtMjdhZmRhMmI0MjA0YzUyMmVlZjctODJh_PF84_8136c735-c7a1-4e2f-98fe-6898007a4ede'

api = CiscoSparkAPI(access_token=BOT_TOKEN)


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
            #sparkMsgText = sparkMsgText.split(botFirstName,1)[1] # Remove bot's first name from message
            sparkMsgPersonEmail = str(sparkMessage.personEmail) # Get message personEmail

            if "hello" in sparkMsgText: #Replies to the Message hello
                textAnswer = 'Hello <@personEmail:' + str(jsonAnswer['data']['personEmail']) + '>, HELLO!!!!!!'
                botAnswered = api.messages.create(roomId=sparkMsgRoomId, markdown=textAnswer)

            elif "help" in sparkMsgText: #Replies to the Message help
                textAnswer = 'Hello <@personEmail:' + str(jsonAnswer['data']['personEmail']) + '>, I will Help you to Invite People to a Spark Space or Team'
                botAnswered = api.messages.create(roomId=sparkMsgRoomId, markdown=textAnswer, files=["https://raw.githubusercontent.com/robertecke83/invitebot2/master/example.png"])

            else:

                 if "@meetingzone.com" in sparkMsgPersonEmail: # Check if the Message comes from a @meetingzone.com domain

                    if not sparkMessage.files: #IF there is no CSV file attached use help
                        textAnswer = 'Hello <@personEmail:' + str(jsonAnswer['data']['personEmail']) + '> I´m missing the *.CSV please @mention me with help to find out how to use me'
                        botAnswered = api.messages.create(roomId=sparkMsgRoomId, markdown=textAnswer)

                        # If the message comes with a file
                    else:
                        sparkMsgFileUrl = str(sparkMessage.files[0]) # Get the URL of the first file
                        
                        sparkHeader = {'Authorization': "Bearer " + BOT_TOKEN}
                        i = 0 # Index to skip title row in the CSV file

                        with requests.Session() as s: # Creating a session to allow several HTTP messages with one TCP connection
                            getResponse = s.get(sparkMsgFileUrl, headers=sparkHeader) # Get file

                            # If the file extension is CSV
                            if str(getResponse.headers['Content-Type']) == 'text/csv':
                                decodedContent = getResponse.content.decode('utf-8')
                                csvFile = csv.reader(decodedContent.splitlines(), delimiter=';')
                                listEmails = list(csvFile)
                                textAnswer = 'Hello <@personEmail:' + str(jsonAnswer['data']['personEmail']) + '>, I will start the Invite now' #Message to Room Invite will start
                                botAnswered = api.messages.create(roomId=sparkMsgRoomId, markdown=textAnswer)
                                for row in listEmails: # Creating one list for each line in the file
                                    if i != 0:
                                        participantAdded = api.memberships.create(roomId=sparkMsgRoomId, personEmail=str(row[2]), isModerator=False) # Add participant from e-mail field
                                    i += 1

                                      
                            else:   # If the attached file is not a CSV
                                textAnswer = 'Sorry, I only understand **CSV** files.'
                                botAnswered = api.messages.create(roomId=sparkMsgRoomId, markdown=textAnswer)

                 else:
                    textAnswer = 'Sorry, Im only allowed to invite people for MeetingZone Employes.'
                    botAnswered = api.messages.create(roomId=sparkMsgRoomId, markdown=textAnswer)


    return 'OK'

if __name__ == '__main__':
    app.run()
