const request = require('request')

require('dotenv').config();

module.exports = {
    getAccessTokenWithRefreshToken: (refreshToken, callback) => {
        var data = 'grant_type=refresh_token'
                + '&refresh_token=' + refreshToken
                + '&client_id=' + process.env.AZUREAD_APP_ID
                + '&client_secret=' + encodeURIComponent(process.env.AZUREAD_APP_PASSWORD)

        var options = {
            method: 'POST',
            url: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
            body: data,
            json: true,
            headers: { 'Content-Type' : 'application/x-www-form-urlencoded' }
        };

        request(options, function (err, res, body) {
            if (err) return callback(err, body, res);
            if (parseInt(res.statusCode / 100, 10) !== 2) {
                if (body.error) {
                    return callback(new Error(res.statusCode + ': ' + (body.error.message || body.error)), body, res);
                }
                if (!body.access_token) {
                    return callback(new Error(res.statusCode + ': refreshToken error'), body, res);
                }
                return callback(null, body, res);
            }
            callback(null, {
                accessToken: body.access_token,
                refreshToken: body.refresh_token
            }, res);
        }); 
    },

    sendWfhEmail: (accessToken, toEmail, sendWFHData, callback) => {
        var data = {
            "Subject": sendWFHData.tbsubject,
            "Body": {
            "ContentType": "HTML",
            "Content": sendWFHData.tbdescription
            },
            "Start": {
                "DateTime": sendWFHData.startdate.toISOString(),
                "TimeZone": "Asia/Singapore"
            },
            "End": {
                "DateTime": sendWFHData.enddate.toISOString(),
                "TimeZone": "Asia/Singapore"
            },
            "IsAllDay": true,
            "ResponseRequested": false,
            "ShowAs": "Free", // https://msdn.microsoft.com/en-us/office/office365/api/complex-types-for-mail-contacts-calendar#FreeBusyStatus
            "IsReminderOn": false,
            "Attendees": [
                {
                    "EmailAddress": {
                    "Address": toEmail,
                    "Name": "My team"
                    },
                    "Type": "Required" // Single instance
                }
            ]
        };

        var options = {
            method: 'POST',
            url: 'https://outlook.office.com/api/v2.0/me/events',
            body: data,
            json: true,
            headers: { 
                'Content-Type' : 'application/json',
                Authorization: 'Bearer ' + accessToken
            }
        };
        request(options, function (err, res, body) {
            if (err) return callback(err, null, null);
            callback(err, body, res);
        }); 
    }
}