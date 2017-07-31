module.exports = () => {
    var api = require('./apiHandler')
    var attachments = require('./attachments')
    bot.dialog('/sendWfhEmail', [
        (session, args, next) => {
            console.log(args)
            var dateRange = getDateRanges(args);
            var wfhCard = {};
            if (dateRange){
                // Person specified dates
                wfhCard = attachments.wfhCardDateTB(dateRange, session.userData.teamEmail, session.userData.accessToken);
            } else {
                wfhCard = attachments.wfhCardDateInput(session.userData.teamEmail, session.userData.accessToken);
            }
             
            var adaptiveCardMessage = new builder.Message(session)
            .addAttachment(wfhCard);
            session.endDialog(adaptiveCardMessage);

            // session.send("Dates: " + dateRange.start + " to " + dateRange.end);
            
            // api.sendWfhEmail(session.userData.accessToken, session.userData.teamEmail, dateRange,
            //     function (err, body, res) {
            //         if (err){
            //             session.endDialog("Hmm something went wrong");
            //         } else if (parseInt(res.statusCode / 100, 10) !== 2) {
            //             // console.log(body);
            //             api.getAccessTokenWithRefreshToken(session.userData.refreshToken, (err, body, res) => {
            //             if (err || body.error) {
            //                 session.send("Error while getting a new access token. Please try logout and login again. Error: " + err);
            //                 session.endDialog();
            //                 }else{
            //                 session.userData.accessToken = body.accessToken;
            //                 console.log(dateRange)
            //                 api.sendWfhEmail(session.userData.accessToken, session.userData.teamEmail, dateRange,
            //                     function (err, body, res) {
            //                     if (parseInt(res.statusCode / 100, 10) !== 2) {
            //                         session.endDialog("Hmm, something went wrong and I couldn't send your wfh message. Sorry.");
            //                     } else {
            //                         session.endDialog("Okay, I've sent a wfh message to the team.");
            //                     }
            //                     });
            //                 }
            //             });
            //         } else {
            //             session.endDialog("Okay, I've sent a wfh message to your team.");
            //         }
            //     });
        }
    ]);
}

function getDateRanges(args) {
    var dateentity = builder.EntityRecognizer.findEntity(args.entities, 'builtin.datetimeV2.date');
    var daterangeentity = builder.EntityRecognizer.findEntity(args.entities, 'builtin.datetimeV2.daterange');
    var dateRange = {};
    if (!isEmptyObject(dateentity)) {
        console.log("in date entity")
        // Start and end date are the same for 1 day WFH
        dateRange.start = new Date(dateentity.resolution.values[0]['value']);
        dateRange.end = new Date(dateentity.resolution.values[0]['value']);
    } else if (!isEmptyObject(daterangeentity)) {
        if (daterangeentity.resolution.values.length > 1) {
            dateRange.start = new Date(daterangeentity.resolution.values[1]['start']);
            dateRange.end = new Date(daterangeentity.resolution.values[1]['end']);
        } else {
            dateRange.start = new Date(daterangeentity.resolution.values[0]['start']);
            dateRange.end = new Date(daterangeentity.resolution.values[0]['end']);
        }
    } else {
        return null
    }
    return dateRange;
}

function isEmptyObject(obj) {
  for (var key in obj) {
    if (Object.prototype.hasOwnProperty.call(obj, key)) {
      return false;
    }
  }
  return true;
}

// Make a separate WFH one
//https://msdn.microsoft.com/en-us/office/office365/api/api-catalog
// function sendWfhEmail(accessToken, callback) {
//   var data = {
//     "Subject": "Alyssa WFH bot test",
//     "Body": {
//       "ContentType": "HTML",
//       "Content": "Working from home"
//     },
//     "Start": {
//         "DateTime": "2017-06-20T00:00:00",
//         "TimeZone": "Asia/Singapore"
//     },
//     "End": {
//         "DateTime": "2017-06-21T00:00:00",
//         "TimeZone": "Asia/Singapore"
//     },
//     "IsAllDay": true,
//     "ResponseRequested": false,
//     "ShowAs": "Free", // https://msdn.microsoft.com/en-us/office/office365/api/complex-types-for-mail-contacts-calendar#FreeBusyStatus
//     "IsReminderOn": false,
//     "Attendees": [
//       {
//         "EmailAddress": {
//           "Address": "ongalyssa@outlook.com",
//           "Name": "Aleessa"
//         },
//         "Type": "Required" // Single instance
//       }
//     ]
//   };

//   var options = {
//       method: 'POST',
//       url: 'https://outlook.office.com/api/v2.0/me/events',
//       body: data,
//       json: true,
//       headers: { 
//         'Content-Type' : 'application/json',
//         Authorization: 'Bearer ' + accessToken
//       }
//   };
//   request(options, function (err, res, body) {
//       if (err) return callback(err, null, null);
//       callback(err, body, res);
//   }); 
// }