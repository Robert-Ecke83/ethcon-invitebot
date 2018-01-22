from flask import Flask, request
from ciscosparkapi import CiscoSparkAPI
import json
import requests
import csv

app = Flask(__name__)

BOT_TOKEN = 'YzkxNThlN2ItZDhiYi00N2UwLWJmNTYtMTExN2I3NDhhZmMxOWQ5MzM4YjUtNzM0'
SPACE_ID = 'Y2lzY29zcGFyazovL3VzL1JPT00vZTJhN2I2ZTAtZmM0My0xMWU3LTljNTEtMjU2ODE2NDAzYzUz'

api = CiscoSparkAPI(access_token=BOT_TOKEN)


@app.route('/')
def hello():
    return 'Hello World!'

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
            sparkMessageID = api.messages.get(jsonAnswer['data']['roomId'])
            sparkMsgText = str(sparkMessage.text) # Get message text
            sparkMsgText = sparkMsgText.split(botFirstName,1)[1] # Remove bot's first name from message

            # Say hello if the message doesn't contain a file
            if not sparkMessage.files:
                textAnswer = 'Hello <@personEmail:' + str(jsonAnswer['data']['personEmail']) + '>, you can send me a CSV file including a list of e-mail addresses and I will add them to this space.'
                botAnswered = api.messages.create(roomId=, markdown=textAnswer)

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
                        for row in listEmails: # Creating one list for each line in the file
                            if i != 0:
                                participantAdded = api.memberships.create(roomId=SPACE_ID, personEmail=str(row[2]), isModerator=False) # Add participant from e-mail field
                                #botAnswered = api.messages.create(roomId=SPACE_ID, text=str(api.exceptions.SparkApiError))
                            i += 1

                    # If the attached file is not a CSV
                    else:
                        textAnswer = 'Sorry, I only understand **CSV** files.'
                        botAnswered = api.messages.create(roomId=SPACE_ID, markdown=textAnswer)


    return 'OK'

if __name__ == '__main__':
    app.run()
