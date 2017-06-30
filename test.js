var builder = require('botbuilder');
var restify = require('restify');
var azure = require('azure-storage');
// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
    console.log('%s listening to %s', server.name, server.url);
});

// Create chat bot and listen to messages
var connector = new builder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});
server.post('/api/messages', connector.listen());

// Bot setup
var bot = new builder.UniversalBot(connector, [
    function (session) {
        session.send('Hi there! Let us start by personalizing your profile. Please choose your persona in this company:');
        var cards = getCardsAttachments();

        // create reply with Carousel AttachmentLayout
        var reply = new builder.Message(session)
            .attachmentLayout(builder.AttachmentLayout.carousel)
            .attachments(cards);

        session.send(reply);
        builder.Prompts.text(session, "Pick one from options 1-3 ");
    },
    function (session, result) {
        var rep = result.response;
        var entGen = azure.TableUtilities.entityGenerator;
        var task = {
            PartitionKey: entGen.String('user'),
            RowKey: entGen.String(rep)
            //description: entGen.String(), store name of user
            // dueDate: entGen.DateTime(new Date(Date.UTC(2015, 6, 20))),
        };
        
        //session.send('creating table');

        //session.send('before insert');
        tableSvc.insertEntity('tablenew', task, function (error, result, response) {
            if (!error) {
                // Entity inserted
                session.send('saved in table');
            }
            else {
                session.send(error);
            }
        });

        session.send('That is great! What would you like to do today?');
        session.send('1. Get introduced to the new workspace - https://ncmedia.azureedge.net/ncmedia/2017/06/MS_Workplace2020_Singapore_EL_office365-1.png');
        session.send('2.	See how you can work better https://www.microsoft.com/singapore/modern-workplace/');
        session.send('Thanks! Great first day! Let us catch up soon.');
        /*   try {
               var ind = getIndex();
               session.send('after');
           }
           catch (error) { session.send(error); }
           session.endDialog();
           */
    }
]);
var azure = require('azure-storage');
var tableSvc = azure.createTableService('msteamsstorage', 'KhLvvKS+f11lHS7t0+VBmuJ00Ha8hh1JadDUaC+g8TQ1UnG6J5HmJPcYbVGl6dEfm4VW/VvPsn1Zb5YfyrNXzA==');
tableSvc.createTableIfNotExists('tablenew', function (error, result, response) {
    if (!error) {
        // Table exists or created

    }
});

function getCardsAttachments(session) {
    return [
        new builder.ThumbnailCard(session)
            .title('1. Creator')
            .subtitle('You are a Creator if -')
            .text('You thrive on inventing new ideas and ways to do things differently, often producing inspiring results.You see problems as opportunities and face them head on, while having some fun with it. Anybody can be a Creator. Roles similar to a Creator include - Designer, Writer, Programmer, Marketing'),

        new builder.ThumbnailCard(session)
            .title('2. Innovator')
            .subtitle('You are an Innovator if -')
            .text('You are a thinker.You constantly strive to reinvent, optimize processes and introduce new methods, ideas, or products.You appreciate fact- based approaches to create breakthrough results. Anybody can be an Innovator. Roles similar to an Innovator include - General Manager, Finance, Sales, Engineer, Analyst'),

        new builder.ThumbnailCard(session)
            .title('3. Collaborator')
            .subtitle('You are a Collaborator if -')
            .text('You believe in sharing ideas.When tasked with a project, you will reach out to someone outside of the team because the natural collaborator knows just whom to ask.You love improving people’s lives and the workplace loves you for it. Anybody can be a Collaborator. Roles similar to a Collaborator include - HR, Marketing, Manager, Communications')
    ];
    session.endDialog();
}



//bot.dialog('index', require('./index'));

/*function getIndex(session) {
    session.send('in here');
    var ind = require('index.js');
}*/