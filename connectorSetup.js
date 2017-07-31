module.exports = function () {
    const path = require('path');
    var restify = require('restify');
    global.builder = require('botbuilder');

    //If testing via the emulator, no need for appId and appPassword. If publishing, enter appId and appPassword here 
    var connector = new builder.ChatConnector({
        appId: process.env.MICROSOFT_APP_ID ? process.env.MICROSOFT_APP_ID : '',
        appPassword: process.env.MICROSOFT_APP_PASSWORD ? process.env.MICROSOFT_APP_PASSWORD : ''
    });

    global.bot = new builder.UniversalBot(connector);

    // Setup Restify Server
    global.server = restify.createServer();
    server.listen(process.env.port || 3978, function () {
        console.log('%s listening to %s', server.name, server.url);
    });
    server.post('/api/messages', connector.listen());
    // bot.use(builder.Middleware.dialogVersion({ version: 0.2, resetCommand: /^reset/i }));
    server.get('/code', restify.serveStatic({
    'directory': path.join(__dirname, 'public'),
    'file': 'code.html'
    }));
}