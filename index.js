
// 'use strict';

require('dotenv').config();

const path = require('path');
const botauth = require('botauth');
const restify = require('restify');
const builder = require('botbuilder');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
// const envx = require("envx");
const expressSession = require('express-session');
const https = require('https');
const request = require('request');
const nodemailer = require('nodemailer');

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
server.use(expressSession({ secret: process.env.BOTAUTH_SECRET, resave: true, saveUninitialized: false }));
// //server.use(passport.initialize());

var ba = new botauth.BotAuthenticator(server, bot, { session: true, baseUrl: 'https://msteamsbot.azurewebsites.net', secret: process.env.BOTAUTH_SECRET, successRedirect: '/code' });

ba.provider("aadv2", (options) => {
    // Use the v2 endpoint (applications configured by apps.dev.microsoft.com)
    // For passport-azure-ad v2.0.0, had to set realm = 'common' to ensure authbot works on azure app service
    let oidStrategyv2 = {
        redirectUrl: options.callbackURL, //  redirect: /botauth/aadv2/callback
        realm: process.env.AZUREAD_APP_REALM,
        clientID: process.env.AZUREAD_APP_ID,
        clientSecret: process.env.AZUREAD_APP_PASSWORD,
        identityMetadata: 'https://login.microsoftonline.com/' + process.env.AZUREAD_APP_REALM + '/v2.0/.well-known/openid-configuration',
        skipUserProfile: false,
        validateIssuer: false,
        //allowHttpForRedirectUrl: true,
        responseType: 'code',
        responseMode: 'query',
        scope: ['profile', 'offline_access', 'User.Read'],
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


// connecting with LUIS
var luisRecognizer = new builder.LuisRecognizer(process.env.LUIS_MODEL_URL || "https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/12046620-4a2d-48e9-87f4-caf6d75c9b2e?subscription-key=7c029b229e924655a57eb8afe6dc990a&verbose=true&timezoneOffset=0&q=");
var intentDialog = new builder.IntentDialog({ recognizers: [luisRecognizer] });
bot.dialog('/', intentDialog);

intentDialog.matches(/\b(hi|hello|hey|howdy|what's up)\b/i, '/signin') //Check for greetings using regex
    .matches(/logout/, "/logout")
 //   .matches(/signin/, "/signin")
   // .matches('beauty', '/beauty') //Check for LUIS intent to get definition
    .matches('wines', '/wines') //Check for LUIS intent to get definition
    .matches('fashion', '/fashion') //Check for LUIS intent to answer why it was introduced
    .matches('about', '/about') //Check for LUIS intent to answer why it was introduced
    .onDefault(builder.DialogAction.send("Sorry, I didn't understand what you said.")); //Default message if all checks fail


bot.dialog('/about', function (session) {
    session.send('DFS ("DFS Group") is a Hong Kong based travel retailer of luxury products. Established in 1960, its network consists of duty-free stores stores located in 17 major airports and 18 downtown Galleria stores,[1][2] as well as resort locations worldwide.');
    session.endDialog();
});

bot.dialog('/beauty', [function (session) {
    session.send('Sure, let\'s learn about Beauty and Fragrances today!');

    var cards = getCardsAttachments();

    // create reply with Carousel AttachmentLayout
    var reply = new builder.Message(session)
        .attachmentLayout(builder.AttachmentLayout.carousel)
        .attachments(cards);

    session.send(reply);
    builder.Prompts.text(session,
        "Enter your choice! Pick one from options 1-2 ");
   
   },
   function(session, result) {
       if (result.response == 1)
       {


       }
       else if (result.response == 2) { }
    }
     
]);


bot.dialog('/fashion', function (session) {
    session.send('');
    session.endDialog();
});


bot.dialog("/logout", (session) => {
    ba.logout(session, "aadv2");
    session.endDialog("logged_out");
});

var username; 
bot.dialog("/signin", [].concat(
    ba.authenticate("aadv2"),
    (session, args, skip) => {
        let user = ba.profile(session, "aadv2");
        session.send('Hello ' + user.displayName + ', welcome to DFS!');
        session.endDialog();
        username = user.displayName;
        session.userData.accessToken = user.accessToken;
        session.userData.refreshToken = user.refreshToken;
      //  session.beginDialog('workPrompt');
        session.beginDialog('persona');
    }
));


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
                        getAccessTokenWithRefreshToken(session.userData.refreshToken, (err, body, res) => {

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
                                    }
                                );
                            }

                        });
                    }
                }
            }
        );
      //  session.beginDialog('persona');
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


function getAccessTokenWithRefreshToken(refreshToken, callback) {
    var data = 'grant_type=refresh_token'
        + '&refresh_token=' + refreshToken
        + '&client_id=' + AZUREAD_APP_ID
        + '&client_secret=' + encodeURIComponent(AZUREAD_APP_PASSWORD)

    var options = {
        method: 'POST',
        url: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        body: data,
        json: true,
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
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
}

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

function getCardsAttachments(session) {
    return [
        new builder.ThumbnailCard(session)
            .title('Estee Lauder Foundation')
            .image('https://www.esteelauder.com/media/export/cms/products/558x768/el_sku_RWAP66_558x768_0.jpg')
            .subtitle('Double Wear Nude')
            .text('BENEFITS \n LIGHTWEIGHT.SPF 30 / ANTI - POLLUTION. 24- HOUR WEAR.LIGHT - TO - MEDIUM COVERAGE'),

        new builder.ThumbnailCard(session)
            .title('Estee Lauder Mascara')
            .image('https://www.esteelauder.com/media/export/cms/products/558x768/el_sku_R8G001_558x768_0.jpg')
           .subtitle('Sumptuous Knockout')
               .text('BENEFITS \nLASHES ARE FANNED OUT.LIFTED.DEFINED')
  ];
    session.endDialog();
}

bot.dialog('persona', [

    function (session) {
        //session.send('entered');
        session.send('Let\'s start by knowing which area you would like to learn about today.');

        builder.Prompts.text(session,
            "1. Beauty & Fragrances \n 2. Fashion & Accessories.");
            session.send("Enter 1 or 2!");

    },
    // function (session,result)
    function (session, result) {
        // session.send('entered 2');
        //  session.send(result);
    //    session.send('before if');
        if (result.response == 1) {
            session.beginDialog('beauty');
        }
        else if (result.response == 2) {
            session.beginDialog('fashion');
        }
        else { session.send("Ooh! Plz pick 1 or 2! Let'\s go again. Say Hi"); }
        session.send('That is great! What would you like to do today?');
        session.send('1. Know about Special offers? - https://www.dfs.com/en/singapore/local-events');
        session.send('2. Read some Trips & Tips? - https://www.dfs.com/en/singapore/trips-and-tips-homepage');

        session.send('You can also ask me more details. Try saying "Tell me about DFS" To Logout, say logout');
        session.beginDialog('/');
    }

    ]);