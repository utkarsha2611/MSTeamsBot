
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

var ba = new botauth.BotAuthenticator(server, bot, { session: true, baseUrl: 'https://mwpwebapp.azurewebsites.net', secret: process.env.BOTAUTH_SECRET, successRedirect: '/code' });

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
    .matches('aboutEvent', '/about') //Check for LUIS intent to get definition
    .matches('askQuesn', '/ask') //Check for LUIS intent to get definition
  //  .matches('cannot', '/cannot') //Check for LUIS intent to answer why it was introduced
  //  .matches('forward', '/forward') //Check for LUIS intent to answer why it was introduced
  //  .matches('guestReg', '/guestRegister') //Check for LUIS intent to answer how to access it
    .matches('modernWP', '/modernWP') //Check for LUIS intent to answer how is it different from SAP 
    .matches('unregister', '/unregister') //Check for LUIS intent to answer what it looks like
    .matches('when', '/when') //Check for LUIS intent to answer how it affects pay
    .matches('where', '/where') //Check for LUIS intent to answer how it affects pay
    .matches('who', '/who') //Check for LUIS intent to answer how it affects pay
 //   .matches('why', '/why') //Check for LUIS intent to answer how it affects pay
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



bot.dialog('/modernWP', function (session) {
    session.send('It’s a new way of working! Watch this video to find out more: https://youtu.be/veLoHcgN7pc');
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
        session.send('Hello ' + user.displayName + ', I\'m AidingAly and I can help you navigate Co-Lab, where the \'Future of Work is Creative\'');
        session.endDialog();
        //  username = user.displayName;
        username = user.displayName;
        session.userData.accessToken = user.accessToken;
        session.userData.refreshToken = user.refreshToken;
        //  session.beginDialog('workPrompt');
        session.beginDialog('persona');
    }
));


/*var azure = require('azure-storage');
var tableSvc = azure.createTableService('msteamsstorage', 'KhLvvKS+f11lHS7t0+VBmuJ00Ha8hh1JadDUaC+g8TQ1UnG6J5HmJPcYbVGl6dEfm4VW/VvPsn1Zb5YfyrNXzA==');
tableSvc.createTableIfNotExists('tablenew', function (error, result, response) {
    if (!error) {
        // Table exists or created

    }
});*/

function getDetails(session)
{ return [
        new builder.HeroCard(session)
        .title('Get introduced to the new workspace')
        .text('https://www.microsoft.com/singapore/modern-workplace/'),
          //  .text('Please proceed to Maker\'s Commons - devices and accessories are on display for interactivity purposes.Enjoy!'),
     
        new builder.HeroCard(session)
            .title('See how you can work better')
            .text('https://www.microsoft.com/singapore/modern-workplace/'),
]}
function getCardsAttachments(session) {
    return [
        new builder.HeroCard(session)
            .title('Device Immersion Experience')
            .subtitle('Free and easy, Touch and Feel')
         
            .images([
                //Using this image: http://imgur.com/a/vl59A
                builder.CardImage.create(session, "https://media.licdn.com/mpr/mpr/shrinknp_200_200/AAEAAQAAAAAAAAzjAAAAJDVjNDRkYzM2LTAzZjctNDUwNi1iNTk2LWI4MGE3ZjFiOTI2Zg.jpg")
            ])
            .text('Please proceed to Maker\'s Commons - devices and accessories are on display for interactivity purposes.Enjoy!'),
          //  .buttons(
           //     builder.CardAction.dialogAction(session, "topnews", null, "Top News")),

        new builder.HeroCard(session)
            .title('Exciting expert Surface Pro device demos')
            .images([
                //Using this image: http://imgur.com/a/vl59A
                builder.CardImage.create(session, "https://d388w23p6r1vqc.cloudfront.net/img/profiles/532/profile_pic.png")
            ])
            .text(' Please proceed to Ideation Hub / Learning'),
        
           // .subtitle('You are an Innovator if - ')
          //  .text('You are a thinker.You constantly strive to reinvent, optimize processes and introduce new methods, ideas, or products.You appreciate fact- based approaches to create breakthrough results. Anybody can be an Innovator. Roles similar to an Innovator include - General Manager, Finance, Sales, Engineer, Analyst'),
        new builder.HeroCard(session)
            .title('Modern Workplace business solutions and applications like Microsoft 365 on the Surface Pro')
           // .subtitle('You are a Collaborator if - ')

        .images([
            //Using this image: http://imgur.com/a/vl59A
            builder.CardImage.create(session, "https://media.licdn.com/mpr/mpr/shrinknp_200_200/AAEAAQAAAAAAAA21AAAAJDQxZWQ5YzRiLWM1YTAtNDFmYS05YjJiLTExNDQ2NTIxN2VlMg.jpg")
            ])
            .text('Please proceed to Ideation Hub / 2'),

        new builder.HeroCard(session)
            .title('Modern Meetings with Modern Devices')
            .images([
                //Using this image: http://imgur.com/a/vl59A
                builder.CardImage.create(session, "https://d388w23p6r1vqc.cloudfront.net/img/profiles/532/profile_pic.png")
            ])
            .text('Please proceed to Ideation Hub / 3 '),

        new builder.HeroCard(session)
            .title('Physical set up of the Modern Workplace ')
            .images([
                //Using this image: http://imgur.com/a/vl59A
                builder.CardImage.create(session, "https://d388w23p6r1vqc.cloudfront.net/img/profiles/532/profile_pic.png")
            ])
            .text('Please approach any of the ambassadors and they will direct a Steelcase spokesperson to aid you in the discussion')


    ];
    session.endDialog();
}




bot.dialog('persona', [

    function (session) {
        //session.send('entered');
        session.send('What can I help you with today?');
        builder.Prompts.text(session, 'Say 1 for Event details \n  2 for Modern Workplace knowhow');

        
     /*   builder.Prompts.text(session,
            "Enter your choice! Pick one from options 1-3 ");*/

    },
    // function (session,result)
    function (session, result) {
        // session.send('entered 2');
        //  session.send(result);
        if (result.response == 1) {
            /*  var username = session.userData.name;
              session.send(username);*/
            /* var rep = result.response;
 
             var entGen = azure.TableUtilities.entityGenerator;
            
                 var task = {
                     PartitionKey: entGen.String(username.toString()),
                     RowKey: entGen.String(rep.toString()),
                     description: entGen.String(ques)
                     //  description: entGen.String(username.toString())// store name of user
                     // dueDate: entGen.DateTime(new Date(Date.UTC(2015, 6, 20))),
                 }*/


            //session.send('creating table');

            //  session.send(task.description);
            /*  tableSvc.insertEntity('tablenew', task, function (error, result, response) {
                  if (!error) {
                      // Entity inserted
                      // session.send('saved in table');
                  }
                  else {
                      session.send('You\'re already registered! See you at the event! :)');
                  }
              }); */

            var cards = getCardsAttachments();

            // create reply with Carousel AttachmentLayout
            var reply = new builder.Message(session)
                .attachmentLayout(builder.AttachmentLayout.carousel)
                .attachments(cards);

           session.send(reply);
        }
        else if (result.response == 2) {
            session.send('That is great! What would you like to know?');

            var card = getDetails();

            // create reply with Carousel AttachmentLayout
            var rep = new builder.Message(session)
                .attachmentLayout(builder.AttachmentLayout.carousel)
                .attachments(card);

            session.send(rep);


            //session.send('1. Get introduced to the new workspace - https://ncmedia.azureedge.net/ncmedia/2017/06/MS_Workplace2020_Singapore_EL_office365-1.png');
           // session.send('2. See how you can work better https://www.microsoft.com/singapore/modern-workplace/');
            session.send('It’s a new way of working! Watch this video to find out more: https://youtu.be/veLoHcgN7pc');
           // session.send('You can also ask me more details about the event. Try saying "What is Modern Workplace?" To Logout, say logout');
            session.beginDialog('/');
        }

        //else { builder.Prompts.text(session, "Invalid entry! Please choose from 1-3 only!"); }
        else { session.send("Invalid entry! Let'\s start again. Say Hi"); }
        session.endDialog();
    
    }]);