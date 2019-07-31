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
      getEmployeeData(email).then(function (response) {
        var employeeData = JSON.parse(response);
        var attachment = new builder.ThumbnailCard()
          .title('Your Donation and Volunteering Metrics:')
          .text(`<b>Total Donations:</b> ${employeeData.employeeDonations[0].totalAmount}<br/>` +
            `<b>Total Corporate Match Donations:</b> ${employeeData.companyDonations[0].totalAmount + employeeData.companyDonations[1].totalAmount + employeeData.companyDonations[2].totalAmount}<br/>` +
            `<b>Total Volunteering Hours:</b> ${employeeData.volunteerParticipations.events.totalHours + employeeData.volunteerParticipations.activities.totalHours + employeeData.volunteerParticipations.npoEvents.totalHours}<br/>`)
          .toAttachment()
        var msg = new builder.Message(session)
          .summary('Your Donation and Volunteering Metrics')
          .attachmentLayout('list') // carousel
          .attachments([
            attachment
          ]);
        session.send(msg);
      });


    } else {
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
          session.endDialog('There is some error');
        }
        else {
          var email = '';
          result.forEach(element => {
            if (element.id == session.message.user.id) {
              email = element.userPrincipalName;
            }
          });
          return email;
        }
      }
    );
    return '';
  }

  function getEmployeeData(email) {
    // determines email
    if (email === 'jj.odell@hacko365.onmicrosoft.com') {
      email = 'Jewell.Willett@yourcause.com';
    } else {
      email = 'Wes.Hendrix@yourcause.com';
    }

    const Http = new XMLHttpRequest();
    var url = `https://api.yourcauseuat.com/v1/employees/${email}/`;
    Http.open("GET", url);
    Http.setRequestHeader("x-bb-Key", "84704ed0-a429-4516-8a9d-fccab0bb49aa");
    Http.setRequestHeader("accept", "application/json");
    var employeeId;
    return new Promise(function (resolve, reject) {
      Http.send();
      Http.onreadystatechange = () => {
        if (Http.readyState === 4) {
          // gets the id from the string (JSON.parse() caused errors)
          var searchTerm = '"affiliateEmployeeId":';
          var response = Http.responseText.substr(Http.responseText.indexOf(searchTerm) + searchTerm.length);
          employeeId = response.substr(0, response.indexOf(','));
          url = `https://api.yourcauseuat.com/v1/employees/${employeeId}/metrics/`
          const Http2 = new XMLHttpRequest();
          Http2.open("GET", url);
          Http2.setRequestHeader("x-bb-Key", "84704ed0-a429-4516-8a9d-fccab0bb49aa");
          Http2.setRequestHeader("accept", "application/json");
          Http2.send();
          Http2.onreadystatechange = () => {
            if (Http2.readyState === 4) {
              if (Http2.status >= 300) {
                reject('Error, status code ' + Http2.status);
              } else {
                resolve(Http2.responseText);
              }
            }
          }
        }
      }
    });

  }
};
