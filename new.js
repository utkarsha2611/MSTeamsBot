require('dotenv').config();

const botauth = require("botauth");

const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
const envx = require("envx");
const expressSession = require('express-session');
const https = require('https');
const request = require('request');
const restify = require('restify')

// Custom modules
require('./connectorSetup.js')();
require('./wfh.js')();
var api = require('./apiHandler.js');

// Environment variable checks

const WEBSITE_HOSTNAME = envx("WEBSITE_HOSTNAME");
const PORT = envx("PORT", 3998);
const BOTAUTH_SECRET = envx("BOTAUTH_SECRET");

//bot application identity
//const MICROSOFT_APP_ID = envx("MICROSOFT_APP_ID");
//const MICROSOFT_APP_PASSWORD = envx("MICROSOFT_APP_PASSWORD");

//oauth details for dropbox
const AZUREAD_APP_ID = envx("AZUREAD_APP_ID");
const AZUREAD_APP_PASSWORD = envx("AZUREAD_APP_PASSWORD");
const AZUREAD_APP_REALM = envx("AZUREAD_APP_REALM");

// Setup Restify Server
// var server = restify.createServer();
// server.listen(process.env.port || process.env.PORT || 3978, function () {
//    console.log('%s listening to %s', server.name, server.url); 
// });

// Create chat connector for communicating with the Bot Framework Service
// var connector = new builder.ChatConnector({
//     appId: process.env.MICROSOFT_APP_ID,
//     appPassword: process.env.MICROSOFT_APP_PASSWORD
// });

// Listen for messages from users 
// server.post('/api/messages', connector.listen());
// server.get('/code', restify.serveStatic({
//   'directory': path.join(__dirname, 'public'),
//   'file': 'code.html'
// }));

// Receive messages from the user and respond by echoing each message back (prefixed with 'You said:')
// global.bot = new builder.UniversalBot(connector);


// LUIS setup
var luisRecognizer = new builder.LuisRecognizer(process.env.LUIS_MODEL_URL);
var intentDialog = new builder.IntentDialog({ recognizers: [luisRecognizer] });

bot.dialog('/', [
    (session, args, next) => {
        console.log(session.message)
        if (!session.userData.teamEmail) {
            // If the user hasn't set up their team's email
            // builder.Prompts.text(session, "What's your team email? I'll use this when you send WFH or overseas emails.");
            session.beginDialog('/getTeamEmail');
        } else if (session.message && session.message.value) {
            // A Card's Submit Action obj was received
            processSubmitAction(session, session.message.value);
            return;
        } else {
            session.replaceDialog('/luisDialog');
        }
    }
]);

bot.dialog('/processSendWFH', [
    (session, sendWFHData) => {
        api.sendWfhEmail(session.userData.accessToken, session.userData.teamEmail, sendWFHData,
            function (err, body, res) {
                if (err) {
                    session.endDialog("Hmm something went wrong");
                } else if (parseInt(res.statusCode / 100, 10) !== 2) {
                    // console.log(body);
                    api.getAccessTokenWithRefreshToken(session.userData.refreshToken, (err, body, res) => {
                        if (err || body.error) {
                            session.send("Error while getting a new access token. Please try logout and login again. Error: " + err);
                            session.endDialog();
                        } else {
                            session.userData.accessToken = body.accessToken;
                            api.sendWfhEmail(session.userData.accessToken, session.userData.teamEmail, sendWFHData,
                                function (err, body, res) {
                                    if (parseInt(res.statusCode / 100, 10) !== 2) {
                                        session.endDialog("Hmm, something went wrong and I couldn't send your wfh message. Sorry.");
                                    } else {
                                        session.endDialog("Okay, I've sent a wfh message to the team.");
                                    }
                                });
                        }
                    });
                } else {
                    session.endDialog("Okay, I've sent a wfh message to your team.");
                }
            });
    }
])

function processSubmitAction(session, value) {
    var defaultErrorMessage = 'Please complete the missing parameters';
    switch (value.type) {
        case 'sendWFH':
            // Search, validate parameters
            if (validateSendWFH(value)) {
                // proceed to search
                session.beginDialog('/processSendWFH', value);
            } else {
                session.send(defaultErrorMessage);
            }
            break;

        default:
            // A form data was received, invalid or incomplete since the previous validation did not pass
            session.send(defaultErrorMessage);
    }
}

function validateSendWFH(sendWFH) {
    if (!sendWFH) {
        return false;
    }

    // Startdate and enddate
    var startdate = Date.parse(sendWFH.startdate);
    var hasStartDate = !isNaN(startdate);
    if (hasStartDate) {
        sendWFH.startdate = new Date(startdate);
    }

    var enddate = Date.parse(sendWFH.enddate);
    var hasEndDate = !isNaN(enddate);
    if (hasEndDate) {
        sendWFH.enddate = new Date(enddate);
    }


    // Subject
    var hasSubject = typeof sendWFH.tbsubject === 'string' && sendWFH.tbsubject.length > 0;

    // Description
    var hasDescription = typeof sendWFH.tbdescription === 'string' && sendWFH.tbdescription.length > 0;

    return hasStartDate && hasEndDate && hasSubject && hasDescription;
}

bot.dialog('/getTeamEmail', [
    (session, args, next) => {
        // If the user hasn't set up their team's email
        builder.Prompts.text(session, "What's your team email? I'll use this when you send WFH or overseas emails.");
    }, (session, args, next) => {
        console.log(args)
        session.userData.teamEmail = args.response;
        session.endDialog("Got it.");
    }
])

bot.dialog('/luisDialog', intentDialog);

intentDialog.matches(/logout/, "/logout")
    .matches(/signin/, "/signin")
    .matches(/wfh/, "/wfhMail")
    .matches(/reset/, "/reset")
    .matches("Workfromhome", "/wfhMail")
    .matches("Overseas", "/logout")
    .onDefault(builder.DialogAction.send("Sorry, I didn't understand what you said."));


//=========================================================
// Auth Setup
//=========================================================

server.use(restify.queryParser());
server.use(restify.bodyParser());
server.use(expressSession({ secret: BOTAUTH_SECRET, resave: true, saveUninitialized: false }));
//server.use(passport.initialize());

var ba = new botauth.BotAuthenticator(server, bot, { session: true, baseUrl: `https://${WEBSITE_HOSTNAME}`, secret: BOTAUTH_SECRET, successRedirect: '/code' });

ba.provider("aadv2", (options) => {
    // Use the v2 endpoint (applications configured by apps.dev.microsoft.com)
    // For passport-azure-ad v2.0.0, had to set realm = 'common' to ensure authbot works on azure app service
    let oidStrategyv2 = {
        redirectUrl: options.callbackURL, //  redirect: /botauth/aadv2/callback
        realm: AZUREAD_APP_REALM,
        clientID: AZUREAD_APP_ID,
        clientSecret: AZUREAD_APP_PASSWORD,
        identityMetadata: 'https://login.microsoftonline.com/' + AZUREAD_APP_REALM + '/v2.0/.well-known/openid-configuration',
        skipUserProfile: false,
        validateIssuer: false,
        //allowHttpForRedirectUrl: true,
        responseType: 'code',
        responseMode: 'query',
        scope: ['email', 'profile', 'offline_access', 'https://outlook.office.com/mail.read', 'https://outlook.office.com/calendars.readwrite'],
        passReqToCallback: true
    };

    let strategy = oidStrategyv2;

    return new OIDCStrategy(strategy,
        (req, iss, sub, profile, accessToken, refreshToken, done) => {
            if (!profile.displayName) {
                return done(new Error("No oid found"), null);
            }
            profile.accessToken = accessToken;
            profile.refreshToken = refreshToken;
            done(null, profile);
        });
});

//=========================================================
// Bots Dialogs
//=========================================================

bot.dialog("/logout", (session) => {
    ba.logout(session, "aadv2");
    session.endDialog("logged_out");
});

bot.dialog("/signin", [].concat(
    ba.authenticate("aadv2"),
    (session, args, skip) => {
        let user = ba.profile(session, "aadv2");
        session.endDialog(user.displayName);
        session.userData.accessToken = user.accessToken;
        session.userData.refreshToken = user.refreshToken;
        session.beginDialog('menu');
    }
));


bot.dialog('menu', [
    (session) => {
        msg = new builder.Message(session)
            .attachments([
                new builder.HeroCard(session)
                    .title('Main Menu')
                    .subtitle('What would you like to do next?')
                    .buttons([
                        builder.CardAction.dialogAction(session, 'wfhMail', null, 'Send WFH message')
                    ])
            ]);
        session.endDialog(msg);
    }
]);

bot.beginDialogAction("wfhMail", '/wfhMail')

bot.dialog("/wfhMail", (session, args, next) => {
    session.replaceDialog("/sendWfhEmail", args);
    session.endDialog();
});

bot.dialog('/reset', (session) => {
    session.userData.teamEmail = null;
    session.endDialog("I've resetted your team email.")
})

bot.dialog('workPrompt', [
    (session) => {
        getUserLatestEmail(session.userData.accessToken,
            function (requestError, result) {
                if (result && result.value && result.value.length > 0) {
                    const responseMessage = 'Your latest email is: "' + result.value[0].Subject + '"';
                    session.send(responseMessage);
                    builder.Prompts.confirm(session, "Retrieve the latest email again?");
                } else {
                    console.log('no user returned');
                    if (requestError) {
                        console.error(requestError);
                        session.send(requestError);
                        // Get a new valid access token with refresh token
                        api.getAccessTokenWithRefreshToken(session.userData.refreshToken, (err, body, res) => {

                            if (err || body.error) {
                                session.send("Error while getting a new access token. Please try logout and login again. Error: " + err);
                                session.endDialog();
                            } else {
                                session.userData.accessToken = body.accessToken;
                                getUserLatestEmail(session.userData.accessToken,
                                    function (requestError, result) {
                                        if (result && result.value && result.value.length > 0) {
                                            const responseMessage = 'Your latest email is: "' + result.value[0].Subject + '"';
                                            session.send(responseMessage);
                                            builder.Prompts.confirm(session, "Retrieve the latest email again?");
                                        }
                                    });
                            }

                        });
                    }
                }
            });
    },
    (session, results) => {
        var prompt = results.response;
        if (prompt) {
            session.replaceDialog('workPrompt');
        } else {
            session.endDialog();
        }
    }
]);


// function getAccessTokenWithRefreshToken(refreshToken, callback){
//   var data = 'grant_type=refresh_token'
//         + '&refresh_token=' + refreshToken
//         + '&client_id=' + AZUREAD_APP_ID
//         + '&client_secret=' + encodeURIComponent(AZUREAD_APP_PASSWORD)

//   var options = {
//       method: 'POST',
//       url: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
//       body: data,
//       json: true,
//       headers: { 'Content-Type' : 'application/x-www-form-urlencoded' }
//   };

//   request(options, function (err, res, body) {
//       if (err) return callback(err, body, res);
//       if (parseInt(res.statusCode / 100, 10) !== 2) {
//           if (body.error) {
//               return callback(new Error(res.statusCode + ': ' + (body.error.message || body.error)), body, res);
//           }
//           if (!body.access_token) {
//               return callback(new Error(res.statusCode + ': refreshToken error'), body, res);
//           }
//           return callback(null, body, res);
//       }
//       callback(null, {
//           accessToken: body.access_token,
//           refreshToken: body.refresh_token
//       }, res);
//   }); 
// }

function getUserLatestEmail(accessToken, callback) {
    var options = {
        host: 'outlook.office.com', //https://outlook.office.com/api/v2.0/me/messages
        path: '/api/v2.0/me/MailFolders/Inbox/messages?$top=1',
        method: 'GET',
        headers: {
            'Content-Type': 'application/json',
            Accept: 'application/json',
            Authorization: 'Bearer ' + accessToken
        }
    };
    https.get(options, function (response) {
        var body = '';
        response.on('data', function (d) {
            body += d;
        });
        response.on('end', function () {
            var error;
            if (response.statusCode === 200) {
                callback(null, JSON.parse(body));
            } else {
                error = new Error();
                error.code = response.statusCode;
                error.message = response.statusMessage;
                // The error body sometimes includes an empty space
                // before the first character, remove it or it causes an error.
                body = body.trim();
                error.innerError = body;
                callback(error, null);
            }
        });
    }).on('error', function (e) {
        callback(e, null);
    });
}