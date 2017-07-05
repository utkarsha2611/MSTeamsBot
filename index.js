
// 'use strict';

// require('dotenv').config();

const path = require('path');
const botauth = require('botauth');
const restify = require('restify');
const builder = require('botbuilder');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
// const envx = require("envx");
const expressSession = require('express-session');
const https = require('https');
const request = require('request');

// const WEBSITE_HOSTNAME = envx("WEBSITE_HOSTNAME");
// const PORT = envx("PORT", 3998);
// const BOTAUTH_SECRET = envx("BOTAUTH_SECRET");
//bot application identity
// const MICROSOFT_APP_ID = envx("MICROSOFT_APP_ID");
// const MICROSOFT_APP_PASSWORD = envx("MICROSOFT_APP_PASSWORD");

//oauth details for dropbox
// const AZUREAD_APP_ID = envx("AZUREAD_APP_ID");
// const AZUREAD_APP_PASSWORD = envx("AZUREAD_APP_PASSWORD");
// const AZUREAD_APP_REALM = envx("AZUREAD_APP_REALM");

//=========================================================
// Bot Setup
//=========================================================

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || 3978, function () {
    console.log('%s listening to %s', server.name, server.url);
});

// Create chat bot
var connector = new builder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});
var bot = new builder.UniversalBot(connector);
server.post('/api/messages', connector.listen());
server.get('/code', restify.serveStatic({
    'directory': path.join(__dirname, 'public'),
    'file': 'code.html'
}));

//=========================================================
// Auth Setup
//=========================================================

server.use(restify.queryParser());
server.use(restify.bodyParser());
// server.use(expressSession({ secret: process.env.BOTAUTH_SECRET, resave: true, saveUninitialized: false }));
// //server.use(passport.initialize());

// var ba = new botauth.BotAuthenticator(server, bot, { session: true, baseUrl: 'https://msteamsbot.azurewebsites.net', secret: process.env.BOTAUTH_SECRET, successRedirect: '/code' });

// ba.provider("aadv2", (options) => {
//     // Use the v2 endpoint (applications configured by apps.dev.microsoft.com)
//     // For passport-azure-ad v2.0.0, had to set realm = 'common' to ensure authbot works on azure app service
//     let oidStrategyv2 = {
//         redirectUrl: options.callbackURL, //  redirect: /botauth/aadv2/callback
//         realm: process.env.AZUREAD_APP_REALM,
//         clientID: process.env.AZUREAD_APP_ID,
//         clientSecret: process.env.AZUREAD_APP_PASSWORD,
//         identityMetadata: 'https://login.microsoftonline.com/' + process.env.AZUREAD_APP_REALM + '/v2.0/.well-known/openid-configuration',
//         skipUserProfile: false,
//         validateIssuer: false,
//         //allowHttpForRedirectUrl: true,
//         responseType: 'code',
//         responseMode: 'query',
//         scope: ['email', 'profile', 'offline_access', 'https://outlook.office.com/mail.read'],
//         passReqToCallback: true
//     };

//     let strategy = oidStrategyv2;

//     return new OIDCStrategy(strategy,
//         (req, iss, sub, profile, accessToken, refreshToken, done) => {
//             if (!profile.displayName) {
//                 return done(new Error("No oid found"), null);
//             }
//             profile.accessToken = accessToken;
//             profile.refreshToken = refreshToken;
//             done(null, profile);
//         });
// });


//=========================================================
// Bots Dialogs
//=========================================================


// connecting with LUIS
var luisRecognizer = new builder.LuisRecognizer(process.env.LUIS_MODEL_URL || "https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/12046620-4a2d-48e9-87f4-caf6d75c9b2e?subscription-key=7c029b229e924655a57eb8afe6dc990a&verbose=true&timezoneOffset=0&q=");
var intentDialog = new builder.IntentDialog({ recognizers: [luisRecognizer] });
bot.dialog('/', intentDialog);

intentDialog.matches(/\b(hi|hello|hey|howdy|what's up)\b/i, '/sayHi') //Check for greetings using regex
    .matches(/logout/, "/logout")
    .matches(/signin/, "/signin")
    .matches('aboutEvent', '/about') //Check for LUIS intent to get definition
    .matches('askQuesn', '/ask') //Check for LUIS intent to get definition
    .matches('cannot', '/cannot') //Check for LUIS intent to answer why it was introduced
    .matches('forward', '/forward') //Check for LUIS intent to answer why it was introduced
    .matches('guestReg', '/guestRegister') //Check for LUIS intent to answer how to access it
    .matches('modernWP', '/modernWP') //Check for LUIS intent to answer how is it different from SAP 
    .matches('unregister', '/unregister') //Check for LUIS intent to answer what it looks like
    .matches('when', '/when') //Check for LUIS intent to answer how it affects pay
    .matches('where', '/where') //Check for LUIS intent to answer how it affects pay
    .matches('who', '/who') //Check for LUIS intent to answer how it affects pay
    .matches('why', '/why') //Check for LUIS intent to answer how it affects pay
    .onDefault(builder.DialogAction.send("Sorry, I didn't understand what you said.")); //Default message if all checks fail


bot.dialog('/about', function (session) {
    session.send('It is a showcase of a Modern Workplace where you can work in a whole new way.You don’t want to miss it.');
    session.endDialog();
});

bot.dialog('/sayHi', function (session) {
    session.send('Try saying things like "What is the Modern workplace?"');
    session.endDialog();
});

bot.dialog('/ask', function (session) {
    session.send('Sure, send me the question and I will get back to you in a bit.');
    session.endDialog();
});

bot.dialog('/cannot', function (session) {
    session.send('Oh no :( Thanks for letting me know, I hope to see you at the next meeting.');
    session.endDialog();
});

bot.dialog('/forward', function (session) {
    session.send('Sure, please extend the invitation to your colleague and have them register.');
    session.endDialog();
});

bot.dialog('/guestRegister', function (session) {
    session.send('You may send the invite to your colleague. The registration link is available in the invite.');
    session.endDialog();
});

bot.dialog('/modernWP', function (session) {
    session.send('It’s a new way of working! Watch this video to find out more: https://youtu.be/veLoHcgN7pc');
    session.endDialog();
});

bot.dialog('/unregister',[
    function (session) {
        session.send('Why? :( I hope you do :)');
    }, 
    function (session) {
        bot.dialog('/cannot', function (session) {
            session.send('Oh no :( Thanks for letting me know, I hope to see you at the next meeting.');
            session.endDialog();
        })}
]    );

bot.dialog('/where', function (session) {
    session.send('Steelcase office, 57 Mohammed Sultan Road');
    session.endDialog();
});

bot.dialog('/when', function (session) {
    session.send('Tuesday, 15 August, 2017');
    session.endDialog();
    });
bot.dialog('/who', function (session) {
    session.send('Let me assist you. Send me the question and I will get back to you in a bit.');
        session.endDialog();
 });
bot.dialog('/why', function (session) {
    session.send();
    session.endDialog();
    });

bot.dialog("/logout", (session) => {
    ba.logout(session, "aadv2");
    session.endDialog("logged_out");
});

// bot.dialog("/signin", [].concat(
//     ba.authenticate("aadv2"),
//     (session, args, skip) => {
//         let user = ba.profile(session, "aadv2");
//         session.endDialog(user.displayName);
//         session.userData.accessToken = user.accessToken;
//         session.userData.refreshToken = user.refreshToken;
//         session.beginDialog('workPrompt');
//     }
// ));

// bot.dialog('workPrompt', [
//     (session) => {
//         getUserLatestEmail(session.userData.accessToken,
//             function (requestError, result) {
//                 if (result && result.value && result.value.length > 0) {
//                     const responseMessage = 'Your latest email is: "' + result.value[0].Subject + '"';
//                     session.send(responseMessage);
//                     builder.Prompts.confirm(session, "Retrieve the latest email again?");
//                 } else {
//                     console.log('no user returned');
//                     if (requestError) {
//                         console.error(requestError);
//                         session.send(requestError);
//                         // Get a new valid access token with refresh token
//                         getAccessTokenWithRefreshToken(session.userData.refreshToken, (err, body, res) => {

//                             if (err || body.error) {
//                                 session.send("Error while getting a new access token. Please try logout and login again. Error: " + err);
//                                 session.endDialog();
//                             } else {
//                                 session.userData.accessToken = body.accessToken;
//                                 getUserLatestEmail(session.userData.accessToken,
//                                     function (requestError, result) {
//                                         if (result && result.value && result.value.length > 0) {
//                                             const responseMessage = 'Your latest email is: "' + result.value[0].Subject + '"';
//                                             session.send(responseMessage);
//                                             builder.Prompts.confirm(session, "Retrieve the latest email again?");
//                                         }
//                                     }
//                                 );
//                             }

//                         });
//                     }
//                 }
//             }
//         );
//     },
//     (session, results) => {
//         var prompt = results.response;
//         if (prompt) {
//             session.replaceDialog('workPrompt');
//         } else {
//             session.endDialog();
//         }
//     }
// ]);


// function getAccessTokenWithRefreshToken(refreshToken, callback) {
//     var data = 'grant_type=refresh_token'
//         + '&refresh_token=' + refreshToken
//         + '&client_id=' + AZUREAD_APP_ID
//         + '&client_secret=' + encodeURIComponent(AZUREAD_APP_PASSWORD)

//     var options = {
//         method: 'POST',
//         url: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
//         body: data,
//         json: true,
//         headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
//     };

//     request(options, function (err, res, body) {
//         if (err) return callback(err, body, res);
//         if (parseInt(res.statusCode / 100, 10) !== 2) {
//             if (body.error) {
//                 return callback(new Error(res.statusCode + ': ' + (body.error.message || body.error)), body, res);
//             }
//             if (!body.access_token) {
//                 return callback(new Error(res.statusCode + ': refreshToken error'), body, res);
//             }
//             return callback(null, body, res);
//         }
//         callback(null, {
//             accessToken: body.access_token,
//             refreshToken: body.refresh_token
//         }, res);
//     });
// }

// function getUserLatestEmail(accessToken, callback) {
//     var options = {
//         host: 'outlook.office.com', //https://outlook.office.com/api/v2.0/me/messages
//         path: '/api/v2.0/me/MailFolders/Inbox/messages?$top=1',
//         method: 'GET',
//         headers: {
//             'Content-Type': 'application/json',
//             Accept: 'application/json',
//             Authorization: 'Bearer ' + accessToken
//         }
//     };
//     https.get(options, function (response) {
//         var body = '';
//         response.on('data', function (d) {
//             body += d;
//         });
//         response.on('end', function () {
//             var error;
//             if (response.statusCode === 200) {
//                 callback(null, JSON.parse(body));
//             } else {
//                 error = new Error();
//                 error.code = response.statusCode;
//                 error.message = response.statusMessage;
//                 // The error body sometimes includes an empty space
//                 // before the first character, remove it or it causes an error.
//                 body = body.trim();
//                 error.innerError = body;
//                 callback(error, null);
//             }
//         });
//     }).on('error', function (e) {
//         callback(e, null);
//     });
// }
