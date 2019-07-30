'use strict';

module.exports.setup = function (app) {
  var builder = require('botbuilder');
  var teams = require('botbuilder-teams');
  var config = require('config');
  var XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
  const https = require('https');

  const subKey = '84704ed0-a429-4516-8a9d-fccab0bb49aa'; // '035f143314da4c2cb6b81542f30639c7';
  const host = 'api.yourcauseuat.com';

  if (!config.has("bot.appId")) {
    // We are running locally; fix up the location of the config directory and re-intialize config
    process.env.NODE_CONFIG_DIR = "../config";
    delete require.cache[require.resolve('config')];
    config = require('config');
  }
  // Create a connector to handle the conversations
  var connector = new teams.TeamsChatConnector({
    // It is a bad idea to store secrets in config files. We try to read the settings from
    // the config file (/config/default.json) OR then environment variables.
    // See node config module (https://www.npmjs.com/package/config) on how to create config files for your Node.js environment.
    appId: config.get("bot.appId"),
    appPassword: config.get("bot.appPassword"),
    authToken: config.get("bot.authToken"),
    envId: config.get("bot.envId"),
    subKey: config.get("bot.subKey")
  });

  var inMemoryBotStorage = new builder.MemoryBotStorage();

  // Define a simple bot with the above connector that echoes what it received
  var bot = new builder.UniversalBot(connector, function (session) {
    // Message might contain @mentions which we would like to strip off in the response
    var text = teams.TeamsMessage.getTextWithoutMentions(session.message);

    var email = getEmail(session);

    if (text === 'me') {
      var response = 'Here is your data: ' + getConstituentData(session);
      session.send(response);
    }
    // if (text === 'grants') {

    //   var attachment1 = new builder.ThumbnailCard()
    //     .title('Kite Foundation')
    //     .text('<b>Deadline:</b> <span style="background-color: #f7a08f">7/15/2019</span><br/>' +
    //       '<b>Funding range:</b> $500 - $5,000<br/>' +
    //       'Accepting Applications')
    //     .toAttachment()

    //   var attachment2 = new builder.ThumbnailCard()
    //     .title('Post & Courier')
    //     .text('<b>Deadline:</b> <span style="background-color: #ffd597">7/25/2019</span><br/>' +
    //       '<b>Funding range:</b> $10,000<br/>' +
    //       'Accepting Applications')
    //     .toAttachment()

    //   var msg = new builder.Message(session)
    //     .summary('Grant applications')
    //     .attachmentLayout('list') // carousel
    //     .attachments([
    //       attachment1,
    //       attachment2
    //     ]);
    //   session.send(msg);

    // }
    else {
      session.send('You said: %s', text);
    }

  }).set('storage', inMemoryBotStorage);

  // Setup an endpoint on the router for the bot to listen.
  // NOTE: This endpoint cannot be changed and must be api/messages
  app.post('/api/messages', connector.listen());

  // Listen for compose messages for linking
  app.post('/api/composeExtension', connector.listen());

  // Export the connector for any downstream integration - e.g. registering a messaging extension
  module.exports.connector = connector;

  function getEmail(session) {
    var conversationId = session.message.address.conversation.id;
    connector.fetchMembers(
      (session.message.address).serviceUrl,
      conversationId,
      (err, result) => {
        if (err) {
          console.log(1);
          // session.endDialog('There is some error');
        }
        else {
          var record = JSON.stringify(result);
          console.log(record);
          console.log(result.userPrincipalName);
          return result.userPrincipalName;
          // session.endDialog('%s', JSON.stringify(result));
        }
      }
    );
  }

  function getConstituentData(text) {
    // var options = {
    //   // headers: {
    //   //   'x-bb-Key': subKey,
    //   //   'accept': 'text/plain'
    //   // },
    //   method: 'GET',
    //   protocol: 'https:',
    //   defaultPort: 443,
    //   host: host,
    //   path: '/v1/employees/Jewell.Willett@yourcause.com/'
    // };
    // console.log('test');
    // const req = https.request(options, (resp) => {

    //   console.log('test2');
    //   let data = ''

    //   resp.on('data', (chunk) => {
    //     console.log('test2');
    //     data += chunk;
    //     console.log(chunk);
    //   });

    //   resp.on('end', () => {
    //     console.log('test2');
    //     console.log(data);
    //   });
    // });
    console.log('test');
    const Http = new XMLHttpRequest();
    const url = 'https://api.yourcauseuat.com/v1/metrics/give';
    Http.open("GET", url);
    Http.setRequestHeader("x-bb-Key", "84704ed0-a429-4516-8a9d-fccab0bb49aa");
    console.log(Http.getAllResponseHeaders());
    Http.send();
    Http.onreadystatechange = (e) => {
      console.log(Http.responseText)
    }
    console.log('here');
    return '';

  }
};
