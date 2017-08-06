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
bot.use(builder.Middleware.dialogVersion({ version: 1.0, resetCommand: /^reset/i }));

//=========================================================
// Bots Dialogs
//=========================================================
// connecting with LUIS
var luisRecognizer = new builder.LuisRecognizer(process.env.LUIS_MODEL_URL || "https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/12046620-4a2d-48e9-87f4-caf6d75c9b2e?subscription-key=7c029b229e924655a57eb8afe6dc990a&verbose=true&timezoneOffset=0&q=");
var intentDialog = new builder.IntentDialog({ recognizers: [luisRecognizer] });
bot.dialog('/', intentDialog);

intentDialog.matches(/\b(hi|hello|hey|howdy|what's up)\b/i, '/sayHi') //Check for greetings using regex
    .matches('aboutEvent', '/about') //Check for LUIS intent to get definition
    .matches('askQuesn', '/ask') //Check for LUIS intent to get definition
    .onDefault(builder.DialogAction.send("Sorry, I didn't understand what you said.")); //Default message if all checks fail

bot.dialog('/sayHi', function (session) {
    session.send('Hello! I\'m AidingAly and I\'m here to help you navigate Co-Lab, where The Future of Work is Creative.');
    session.send('To help you get the most out of Co-Lab, I will be here at all times to help you navigate the event.');
    session.endDialog();
    session.beginDialog('/persona');
});

function getDetails(session) {
    return [
        new builder.HeroCard(session)
            .title('Get introduced to the new workspace')
            .buttons([
                builder.CardAction.openUrl(session, 'https://www.microsoft.com/singapore/modern-workplace/', 'Learn More')
            ]),
        new builder.HeroCard(session)
            .title('See how you can work better')
            .buttons([
                builder.CardAction.openUrl(session, 'https://www.microsoft.com/singapore/modern-workplace/', 'Learn More')
            ]),
        new builder.HeroCard(session)
            .title('Watch Modern Workplace video')
            .buttons([
                builder.CardAction.openUrl(session, 'https://youtu.be/veLoHcgN7pc', 'Watch Video')
            ])
    ]
}
function getCardsAttachments(session) {
    return [
        new builder.HeroCard(session)
            .title('Here\'s the site map for your reference')
            .images([
                builder.CardImage.create(session, "https://github.com/utkarsha2611/MSTeamsBot/blob/master/images/GetImage.jpg")
            ]),
        new builder.HeroCard(session)
            .title('1: Device Interactivity - An immersive device experience')
            .images([
                //Using this image: http://imgur.com/a/vl59A
                builder.CardImage.create(session, "https://github.com/utkarsha2611/MSTeamsBot/blob/master/images/makerscommons.PNG")
            ])
            .text('Please proceed to <b>Makers Commons</b> - devices and accessories are on display for interactivity purposes. Enjoy!'),

        new builder.HeroCard(session)
            .title('2: Exciting expert Surface Pro device demos')
            .images([
                //Using this image: http://imgur.com/a/vl59A
                builder.CardImage.create(session, "https://d388w23p6r1vqc.cloudfront.net/img/profiles/532/profile_pic.png")
            ])
            .text('Please proceed to: <b>Ideation Hub (Learning) - Modern Devices for Intelligent Integration</b>'),

        new builder.HeroCard(session)
            .title('3: Modern Workplace business solutions and applications like Microsoft 365 on the Surface Pro')
            .images([
                //Using this image: http://imgur.com/a/vl59A
                builder.CardImage.create(session, "https://media.licdn.com/mpr/mpr/shrinknp_200_200/AAEAAQAAAAAAAA21AAAAJDQxZWQ5YzRiLWM1YTAtNDFmYS05YjJiLTExNDQ2NTIxN2VlMg.jpg")
            ])
            .text('Please proceed to:<b> Ideation Hub - Modern Business Solutions</b>'),

        new builder.HeroCard(session)
            .title('4: Modern Meetings with Modern Devices')
            .images([

                builder.CardImage.create(session, "https://d388w23p6r1vqc.cloudfront.net/img/profiles/532/profile_pic.png")
            ])
            .text('Please proceed to :<b> Ideation Hub - Modern Meetings</b>'),

        new builder.HeroCard(session)
            .title('5: Modern Project Management tools for delivering high quality outcomes')
            .images([
                //Using this image: http://imgur.com/a/vl59A
                builder.CardImage.create(session, "https://d388w23p6r1vqc.cloudfront.net/img/profiles/532/profile_pic.png")
            ])
            .text('Please proceed to : <b>Focus Seat - Modern Project Management</b>'),
        new builder.HeroCard(session)
            .title('6: Physical set up of the Modern Workplace')
            .images([
                //Using this image: http://imgur.com/a/vl59A
                builder.CardImage.create(session, "https://d388w23p6r1vqc.cloudfront.net/img/profiles/532/profile_pic.png")
            ])
            .text('Please approach any of the <b>ambassadors</b> and they will direct a Steelcase representative to assist.'),
    ];

    session.endDialog();
}

bot.dialog('/persona', [
    function (session) {
        //session.send('entered');
        session.send('What can I help you with today?');
        builder.Prompts.text(session, 'Say 1 for Event details \n\n  2 for Modern Workplace knowhow');
    },
    // function (session,result)
    function (session, results) {
        if (results.response == 1) {
            session.replaceDialog('/option1');
        }
        else if (results.response == 2) {
            session.replaceDialog('/option2');
        }
        //else { builder.Prompts.text(session, "Invalid entry! Please choose from 1-3 only!"); }
        else { session.endDialog("Invalid entry! Let'\s start again. Say Hi"); }
    }]);

bot.dialog('/option1', [
    function (session) {
        var cards = getCardsAttachments();
        // create reply with Carousel AttachmentLayout
        var reply = new builder.Message(session)
            .attachmentLayout(builder.AttachmentLayout.carousel)
            .attachments(cards);
        session.send(reply);
        builder.Prompts.text(session, 'Hope you\'ve figured out which room to go! Do ping me if you\'d like to know about Modern Workplace. Just say 2');
    },
    function (session, results) {
        if (results.response == 2) {
            session.replaceDialog('/option2');
        }
        else {
            session.send('Alright! If you have any questions, the team from Microsoft is here to help, please approach our Microsoft ambassadors who will help you out. Enjoy yourself at the event! :)');
            session.replaceDialog('/');
        }
    }
]);

bot.dialog('/option2', [
    function (session) {
        session.send('That is great! What would you like to know?');
        var card = getDetails();
        // create reply with Carousel AttachmentLayout
        var rep = new builder.Message(session)
            .attachmentLayout(builder.AttachmentLayout.carousel)
            .attachments(card);
        session.send(rep);
        builder.Prompts.text(session, 'If you would also like to check the event details again, just say 1');
    },
    function (session, results) {
        if (results.response == 1) {
            session.replaceDialog('/option1');
        }
        else {
            session.send('Alright! If you have any questions, the team from Microsoft is here to help, please approach our Microsoft ambassadors who will help you out. Enjoy yourself at the event! :)');
            session.replaceDialog('/');
        }
    }
]);