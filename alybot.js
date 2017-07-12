require('dotenv-extended').load();
const expressSession = require('express-session');
const https = require('https');
const request = require('request');
var needle = require('needle'),
    url = require('url'),
    validUrl = require('valid-url'),
   // captionService = require('./caption-service'),
    restify = require('restify'),
    builder = require('botbuilder'),
    
    translator = require('mstranslator');
const botauth = require('botauth');


//Keys
var key = {
    msTranslator: "39cf6fcb295b4df7b5353c907b6bb2ca", //ToDo - Move to env
    luisAppKey: '56887e19-48dc-4610-be3b-94684a62e',
    luisSubsKey: '7c029b229e924655a57eb8afe6dc990a'
};

//Setup Translator
var client = new translator({
    api_key: key.msTranslator // use this for the new token API. 
}, true);

//Standard Replies
var standardReplies = {
    firstInit: "So, we're meeting for the first time. Do you want to set a preferred communication language for our chats?",
    langChanged: "Okay! Your language preference has been changed. ",
    askQn: "To begin, just ask me a question like",
    queryExample: "\n-Which course do you want to study today?",
   // techLimitation: "Due to technical limitations, please send me your requests in English.",
    didNotUnderstand: "Sorry, I didn't understand what you said.",
    langReset: "Your language preferences have been reset.",
    resetSuccess: "I've just reset myself, lets try again!",
    startCommand: "Hello! I'm Aly.\n\nWhat do you want to do?\n\nEg. ",
    notReady: "Okay! Talk to me when you are ready :)",
    configureLang: "Please configure your language first by saying 'Configure Language'"
};

//=========================================================
// Bot Setup
//=========================================================

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
    console.log('%s listening to %s', server.name, server.url);
});

//LUIS Setup
var recognizer = new builder.LuisRecognizer('https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/' + key.luisAppKey + '?subscription-key=' + key.luisSubsKey + '&verbose=true&timezoneOffset=0&q=');
var intentDialog = new builder.IntentDialog({ recognizers: [recognizer] });

// Create chat bot
var connector = new builder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});
var bot = new builder.UniversalBot(connector, { persistConversationData: true });
server.post('/api/messages', connector.listen());

// Bot Dialogs
bot.dialog('/', intentDialog);

//LUIS Intents
intentDialog.matches(/\b(hello|hi|hey|how are you)\b/i, '/conversation')
    .matches(/\b(rlang)\b/i, '/resetLang')
    .matches(/\b(rbot)\b/i, '/resetLang')
    .onDefault('/defaultResp');

bot.dialog('/defaultResp', function (session, args) {
    if (session.userData['Lang']) {
        var paramsTranslateToDefault = {
            text: standardReplies.didNotUnderstand,
            from: 'en',
            to: session.userData['Lang']
        };
        client.translate(paramsTranslateToDefault, function (err, dataDefault) {
            session.endDialog(dataDefault);
        })
    } else {
        session.endDialog(standardReplies.didNotUnderstand);
    };

});

bot.dialog('/resetLang', function (session, args) {
    session.sendTyping();
    session.userData['Lang'] = null;
    session.endDialog(standardReplies.langReset);
});

bot.dialog('/resetBot', function (session, args) {
    session.sendTyping();
    session.userData['Lang'] = null;
    session.endDialog(standardReplies.resetSuccess);
    session.endConversation();
});

bot.dialog('/conversation', [function (session, args) {
    session.sendTyping();
    var lang = session.userData['Lang'];
    if (!lang) {
        session.send("Hello! I'm Aly! You can shoot your questions at me! Happy to help! :) ");
        console.log("User's first load or language reset.");

        builder.Prompts.choice(session, standardReplies.firstInit, "English|Chinese|Japanese|Tamil|Hindi|Cancel");
    } else {
        console.log("User Lang Data Exists: " + lang);

        if (lang == "en") {
            session.endDialog(standardReplies.startCommand + standardReplies.queryExample);
        } else {
            var paramsTranslateTo = {
                text: standardReplies.startCommand,
                from: 'en',
                to: lang
            };

            client.translate(paramsTranslateTo, function (err, data) {
                session.endDialog(data + standardReplies.queryExample);
            });
        }
    }
}, function (session, results) {
    session.sendTyping();
    if (results.response && results.response.entity !== 'Cancel') {
        var fullLang = "English";

        //Add more languages for your liking, add prompt on top also
        if (results.response.entity.toUpperCase() == "ENGLISH") {
            session.userData['Lang'] = 'en';
            fullLang = "English";
        } else if (results.response.entity.toUpperCase() == "CHINESE") {
            session.userData['Lang'] = 'zh-chs';
            fullLang = "中文";
        } else if (results.response.entity.toUpperCase() == "JAPANESE") {
            session.userData['Lang'] = 'ja';
            fullLang = "日本語";
        } else if (results.response.entity.toUpperCase() == "TAMIL") {
            session.userData['Lang'] = 'ta';
            fullLang = "தமிழ்";
        } else if (results.response.entity.toUpperCase() == "HINDI") {
            session.userData['Lang'] = 'hi';
            fullLang = "हिन्दी";
        }

        var paramsTranslateTo = {
            text: standardReplies.langChanged,
            from: 'en',
            to: session.userData['Lang']
        };

        var paramsTranslateTo2 = {
            text: standardReplies.techLimitation,
            from: 'en',
            to: session.userData['Lang']
        };

        var paramsTranslateTo3 = {
            text: standardReplies.askQn,
            from: 'en',
            to: session.userData['Lang']
        };



        client.translate(paramsTranslateTo, function (err, data) {
            client.translate(paramsTranslateTo2, function (err, data2) {
                client.translate(paramsTranslateTo3, function (err, data3) {
                    session.send(data + "(" + fullLang + ")");
                    session.send(data2);
                    session.endDialog(data3 + standardReplies.queryExample);
                });
            });
        });
    } else {
        session.endDialog(standardReplies.notReady);
    }

    session.beginDialog('department');
}
]);


//Image Caption

/*bot.dialog('/landmark', function (session, args) {
    if (hasImageAttachment(session)) {
        var stream = getImageStreamFromMessage(session.message);
        captionService
            .getCaptionFromStream(stream)
            .then(function (caption) { handleSuccessResponse(session, caption); })
            .catch(function (error) { handleErrorResponse(session, error); });
    }
    else if (session.message.text == 'exit') {
        session.endDialog('Heading Back');
    }
    else {
        var imageUrl = parseAnchorTag(session.message.text) || (validUrl.isUri(session.message.text) ? session.message.text : null);
        if (imageUrl) {
            captionService
                .getCaptionFromUrl(imageUrl)
                .then(function (caption) { handleSuccessResponse(session, caption); })
                .catch(function (error) { handleErrorResponse(session, error); });
        } else {
            session.send('Please upload an image! Try sending an image or an image URL');
        }
    }
});*/

//=========================================================
// Utilities
//=========================================================
/*function hasImageAttachment(session) {
    return session.message.attachments.length > 0 &&
        session.message.attachments[0].contentType.indexOf('image') !== -1;
}

function getImageStreamFromMessage(message) {
    var headers = {};
    var attachment = message.attachments[0];
    if (checkRequiresToken(message)) {
        // The Skype attachment URLs are secured by JwtToken,
        // you should set the JwtToken of your bot as the authorization header for the GET request your bot initiates to fetch the image.
        // https://github.com/Microsoft/BotBuilder/issues/662
        connector.getAccessToken(function (error, token) {
            var tok = token;
            headers['Authorization'] = 'Bearer ' + token;
            headers['Content-Type'] = 'application/octet-stream';

            return needle.get(attachment.contentUrl, { headers: headers });
        });
    }

    headers['Content-Type'] = attachment.contentType;
    return needle.get(attachment.contentUrl, { headers: headers });
}

function checkRequiresToken(message) {
    return message.source === 'skype' || message.source === 'msteams';
}

/**
 * Gets the href value in an anchor element.
 * Skype transforms raw urls to html. Here we extract the href value from the url
 * @param {string} input Anchor Tag
 * @return {string} Url matched or null
 */
/*function parseAnchorTag(input) {
    var match = input.match('^<a href=\"([^\"]*)\">[^<]*</a>$');
    if (match && match[1]) {
        return match[1];
    }

    return null;
}

//=========================================================
// Response Handling
//=========================================================
function handleSuccessResponse(session, caption) {
    if (caption) {
        session.send('I think it\'s ' + caption);
    }
    else {
        session.send('Couldn\'t find a caption for this one');
    }

}

function handleErrorResponse(session, error) {
    var clientErrorMessage = 'Oops! Something went wrong. Try again later.';
    if (error.message && error.message.indexOf('Access denied') > -1) {
        clientErrorMessage += "\n" + error.message;
    }

    console.error(error);
    session.send(clientErrorMessage);
}*/


//=========================================================
// Auth Setup
//=========================================================

server.use(restify.queryParser());
server.use(restify.bodyParser());
server.use(expressSession({ secret: 'msdfsbotsecret', resave: true, saveUninitialized: false }));
// //server.use(passport.initialize());

var ba = new botauth.BotAuthenticator(server, bot, { session: true, baseUrl: 'https://msteamsbot.azurewebsites.net', secret: 'msdfsbotsecret', successRedirect: '/code' });

ba.provider("aadv2", (options) => {
    // Use the v2 endpoint (applications configured by apps.dev.microsoft.com)
    // For passport-azure-ad v2.0.0, had to set realm = 'common' to ensure authbot works on azure app service
    let oidStrategyv2 = {
        redirectUrl: options.callbackURL, //  redirect: /botauth/aadv2/callback
        realm: 'common',
        clientID: 'dee34b4f-11c7-4794-80ad-9c9654047351',
        clientSecret: 'FowbP8R0DJihnRX94tgv7NF',
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


var username;
bot.dialog("/signin", [].concat(
    ba.authenticate("aadv2"),
    (session, args, skip) => {
        let user = ba.profile(session, "aadv2");
        session.send('Hello ' + user.displayName + ', welcome to the organization');
        session.endDialog();
        username = user.displayName;
        session.userData.accessToken = user.accessToken;
        session.userData.refreshToken = user.refreshToken;
        //  session.beginDialog('workPrompt');
       session.beginDialog('conversation');
    }
));


function getCardsAttachments(session) {
    return [
        new builder.ThumbnailCard(session)
            .title('1.Beauty & Fragrances')
            .image('https://www.google.com.sg/search?q=wines+and+spirits&source=lnms&tbm=isch&sa=X&ved=0ahUKEwi56IDRlITVAhWLQo8KHSpHCbAQ_AUICygC&biw=1440&bih=855#imgrc=AsPikwQWA7rm_M:')
           // .subtitle('You are a Creator if - ')
         .text('You thrive on inventing new ideas and ways to do things differently, often producing inspiring results.You see problems as opportunities and face them head on, while having some fun with it. Anybody can be a Creator. Roles similar to a Creator include - Designer, Writer, Programmer, Marketing'),

        new builder.ThumbnailCard(session)
            .title('2. Wines and spirits')
            .subtitle('You are an Innovator if - ')
            .text('You are a thinker.You constantly strive to reinvent, optimize processes and introduce new methods, ideas, or products.You appreciate fact- based approaches to create breakthrough results. Anybody can be an Innovator. Roles similar to an Innovator include - General Manager, Finance, Sales, Engineer, Analyst'),

        new builder.ThumbnailCard(session)
            .title('3. Food & Gifts')
            .subtitle(' ')
            .text('You believe in sharing ideas.When tasked with a project, you will reach out to someone outside of the team because the natural collaborator knows just whom to ask.You love improving people\m’s lives and the workplace loves you for it. Anybody can be a Collaborator. Roles similar to a Collaborator include - HR, Marketing, Manager, Communications')
    ];
    session.endDialog();
}


bot.dialog('department', [

    function (session) {
        //session.send('entered');
        session.send("Let's start by selecting your department :");

        var cards = getCardsAttachments();

        // create reply with Carousel AttachmentLayout
        var reply = new builder.Message(session)
            .attachmentLayout(builder.AttachmentLayout.carousel)
            .attachments(cards);

        session.send(reply);
        builder.Prompts.text(session,
            "Enter your choice! Pick one from options 1-3 ");

    },
    // function (session,result)
    function (session, result) {
        // session.send('entered 2');
        //  session.send(result);
        if (result.response == 1 || result.response == 2 || result.response == 3) {
            /*  var username = session.userData.name;
              session.send(username);*/
            var rep = result.response;
          
            session.send('That is great! What would you like to do today?');
            session.send('1. Read travel journals at DFS? - https://www.dfs.com/en/singapore/trips-and-tips-homepage');
            session.send('2. Check out special offers at DFS? ;) https://www.dfs.com/en/singapore/local-events');

            session.send('You can also ask me more questions. Try saying "What wine should I sell?" To Logout, say logout');
            session.beginDialog('/');

        }
        //else { builder.Prompts.text(session, "Invalid entry! Please choose from 1-3 only!"); }
        else { session.send("That's an invalid option, dear! Let'\s start again. Say Hi"); }


    }]);