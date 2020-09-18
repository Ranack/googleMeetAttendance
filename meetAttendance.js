function listUpcomingEvents() {
  var date = new Date();
  var timezoneOffset = date.getMinutes() + date.getTimezoneOffset();
  var timestamp = date.getTime() + timezoneOffset * 1000;
  var startHour = new Date(timestamp);
  startHour.setUTCHours(4, 0, 0, 0);
  Logger.log(startHour);
  var endHour = new Date(timestamp);
  endHour.setUTCHours(16, 0, 0, 0);
  Logger.log(endHour);
  var calendarId = ''; // mettre id de l'agenda 
  var optionalArgs = {
    timeMin: (new Date(startHour)).toISOString(),
    timeMax: (new Date(endHour)).toISOString(),
    showDeleted: false,
    singleEvents: true,
    orderBy: 'startTime'
  };
  var response = Calendar.Events.list(calendarId, optionalArgs);
  var events = response.items;
  var tab = [response.items];
  if (events.length > 0) {
    for (i = 0; i < events.length; i++) {
      var event = events[i];
      var when = event.start.dateTime;
      if (!when) {
        when = event.start.date;
      }
  var meetCode= [event.conferenceData];
  const eventTest = response["items"][i]
  const conferenceData = eventTest["conferenceData"]
  var codeMeet = conferenceData.conferenceId
  var cleanCode = codeMeet.toUpperCase();
  var goodCode = cleanCode.replace("-","");
  var codeFinal = goodCode.replace("-","");
  Logger.log(codeFinal)
  var nombre = events.length;
  var formSheet = SpreadsheetApp.getActive().getSheetByName("Rapport")
  for(var e=0; e<nombre; i++){
    formSheet.getRange(e+1, 1).setValue();//date et heure de l'inscription
    formSheet.getRange(e+1, 2).setValue();//mail de reception du rapport
    formSheet.getRange(e+1, 3).setValue();//Code Meet
    formSheet.getRange(e+1, 4).setValue();//type de meet
  }
  
    }
  } else {
    Logger.log('No upcoming events found.');
  }
}


function getMeetAttendance() {

    var formSheet = SpreadsheetApp.getActive().getSheetByName("Rapport")
    var data = formSheet.getDataRange().getValues()


    for (i = 1; i < data.length; i++) {
        var userClaimedorganizer = data[i][1]
        var meetingCode = data[i][2].replace(/-/g, "") 
        var meetingType = null
        var status = data[i][4]
        var userEmail = data[i][1]



        if (data[i][3] == "Regular Meeting") {
            meetingType = "call_ended"
            var fileArray = [
                ["Étudiant", "Temps de suivi (en mins)","Retard"]
            ]
        } else {
            meetingType = "livestream_watched"
            var fileArray = [
                ["Étudiant", "Statut de suivi"]
            ]

        }




        if (status != "Report Sent" && status != "User Not Authorized") {
            let activities = new Map();
            var applicationName = "meet"
            var pageToken;
            var date = new Date();
            var timezoneOffset = date.getMinutes() + date.getTimezoneOffset();
            var timestamp = date.getTime() + timezoneOffset * 1000;
            var correctDate = new Date(timestamp);
            correctDate.setUTCHours(3, 0, 0, 0);
            console.log(correctDate);
            var isoDate = correctDate.toISOString();
            console.log(isoDate);

          
          var optionalArgs = {
                event_name: meetingType,
                startTime: isoDate,
                filters: "meeting_code==" + meetingCode,
                pageToken: pageToken
            };

            do {
                apiCall = getMeetingDetails(applicationName, optionalArgs)
                var apiResponse = apiCall.items
                var pageToken = apiCall.nextPageToken
                for (var key in apiResponse) {
                    var events = apiResponse[key]["events"]

                    var apiClaimedorganizer = null



                    events.forEach(function(item) {
                        var parameters = item["parameters"]
                        var obj = {}
                        

   
                        parameters.forEach(function(filter) {

                            if (filter["name"] == "identifier") {
                                obj.email = filter["value"]
                            }

                            if (filter["name"] == "duration_seconds") {
                                obj.duration = Math.abs(filter["intValue"])

                            }

                            if (apiClaimedorganizer == null && filter["name"] == "organizer_email") {
                                apiClaimedorganizer = filter["value"]

                            }
                        })


                        if ((obj.email != "" || obj.email != null)) {

                            if (meetingType == "livestream_watched") {
                                activities.set(obj.email, "Livestream Watched")

                            } else if (meetingType == "call_ended" && obj.duration != null && obj.duration > 0) {
                                if (activities.has(obj.email)) {
                                    activities.set(obj.email, (activities.get(obj.email) + Math.round(obj.duration / 60)))
                                } else {
                                    activities.set(obj.email, Math.round(obj.duration / 60))
                                }
                            }
                        }

                    })
                }

                pageToken = apiCall.nextPageToken

            } while (pageToken)


            for (let [email, activity] of activities) {

                fileArray.push([email, activity])

            }


            var doc = createDocument(meetingCode, fileArray, userEmail)

            if (userClaimedorganizer == userClaimedorganizer) {

                authorizedResponse(formSheet, doc.getUrl(), meetingCode, userEmail, i)

                shareDoc(doc.getId(), userEmail)

            } else {
                unAuthorizedResponse(formSheet, doc.getUrl(), meetingCode, userEmail, i)
            }

        } else {
            //console.log("error")
        }


    }

}





function getMeetingDetails(applicationName, optionalArgs) {
    var apiCall = AdminReports.Activities.list("all", applicationName, optionalArgs)
    return apiCall

}



function authorizedResponse(formSheet, docLink, meetingCode, userEmail, lineNumber) {
    formSheet.getRange(lineNumber + 1, 5).setValue("Report Sent")
    formSheet.getRange(lineNumber + 1, 6).setValue(docLink)
    GmailApp.sendEmail(userEmail, "Rapport présence " + meetingCode, `Voici le rapport de présence pour le cours ${meetingCode}
                       ${docLink}

Cordialement,

Pour toutes questions, remarques, ou questions merci de contacter votre responsable SI.`)

}




function unAuthorizedResponse(formSheet, docLink, meetingCode, userEmail, lineNumber) {
    formSheet.getRange(lineNumber + 1, 5).setValue("erreur authent")
    formSheet.getRange(lineNumber + 1, 6).setValue(docLink)
    GmailApp.sendEmail(userEmail, "Rapport " + meetingCode, `Vous ne disposez pas des droits pour voir ce rapport, merci de contacter votre responsable SI.`)

}



function createDocument(meetingCode, fileArray, userEmail) {
    var style = {};
    style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
        DocumentApp.HorizontalAlignment.RIGHT;
    style[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
    style[DocumentApp.Attribute.FONT_SIZE] = 14;
    style[DocumentApp.Attribute.BOLD] = true;
    var doc = DocumentApp.create("Rapport de présence " + meetingCode)
    doc.getBody().appendParagraph("Rapport détaillé :  " + meetingCode).setHeading(DocumentApp.ParagraphHeading.HEADING1)
    var docId = doc.getId()
    var docLink = doc.getUrl()
    var format = doc.getBody().appendTable(fileArray).getRow(0).setAttributes(style)
    return doc
}


function shareDoc(docId, userEmail) {
    DocumentApp.openById(docId).addViewer(userEmail)
    return userEmail
}