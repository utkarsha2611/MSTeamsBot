// JavaScript source code


const appInsights = require("applicationinsights");
appInsights.setup("795df439-94fb-4fbe-91be-e35d774d1310");
appInsights.start();
var restify = require('restify');
var builder = require('botbuilder');


//=========================================================
// Bot Setup
//=========================================================

// Setup Restify Server
// Listen for any activity on port 3978 of our local server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
    console.log('%s listening to %s', server.name, server.url);
});

// Create chat bot
var connector = new builder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});
var bot = new builder.UniversalBot(connector);
// If a Post request is made to /api/messages on port 3978 of our local server, then we pass it to the bot connector to handle
server.post('/api/messages', connector.listen());

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
    session.send('Hi there! Try saying things like "What is the Modern workplace?"');
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

bot.dialog('/unregister', [
    function (session) {
        session.send('Why? :( I hope you do :)');
        session.endDialog();
         },
    function(session) {
        //bot.dialog('/cannot', function (session) {
        session.send('Oh no :( Thanks for letting me know, I hope to see you at the next meeting.');
        session.endDialog();
        }
    
]);
bot.dialog('/when', function (session) {
    session.send('Tuesday, 15 August, 2017');
    session.endDialog();
});
bot.dialog('/where', function (session) {
    session.send('Steelcase office, 57 Mohammed Sultan Road');
    session.endDialog();
});

bot.dialog('/who', function (session) {
    session.send('Let me assist you. Send me the question and I will get back to you in a bit.');
    session.endDialog();
});

bot.dialog('/why', function (session) {
    session.send('There is a new way of working, dont you want to find out how? :)');
    session.endDialog();
});


exports.createTelemetry = function (session, properties) {
    var data = {
        conversationData: JSON.stringify(session.conversationData),
        privateConversationData: JSON.stringify(session.privateConversationData),
        userData: JSON.stringify(session.userData),
        conversationId: session.message.address.conversation.id,
        userId: session.message.address.user.id
    };

    if (properties) {
        for (property in properties) {
            data[property] = properties[property];
        }
    }

    return data;
};