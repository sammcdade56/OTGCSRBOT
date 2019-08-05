'use strict';

module.exports.setup = function (app) {
    var builder = require('botbuilder');
    var teams = require('botbuilder-teams');
    var config = require('config');

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

        //This is where it tests for which command that was sent
        if (text === 'company') {
            var promise1 = getUserDataWithPromise('https://api.yourcauseuat.com/v1/metrics/give');
            var promise2 = getUserDataWithPromise('https://api.yourcauseuat.com/v1/metrics/volunteer');
            //This waits for both promises to resolve and then parses the data to return 
            Promise.all([promise1, promise2]).then(function (response) {
                var givingData = JSON.parse(response[0]);
                var volunteeringData = JSON.parse(response[1]);
                var totalAmount = 0;
                var totalHours = 0;
                var totalParticipants = 0;
                if (volunteeringData.events != null) {
                    totalHours += volunteeringData.events.totalHours;
                    totalParticipants += volunteeringData.events.totalParticipants;
                }
                if (volunteeringData.activities != null) {
                    totalHours += volunteeringData.activities.totalHours;
                    totalParticipants += volunteeringData.activities.totalParticipants;
                }
                if (volunteeringData.npoEvents != null) {
                    totalHours += volunteeringData.npoEvents.totalHours;
                    totalParticipants += volunteeringData.npoEvents.totalParticipants;
                }

                for (var i = 0; i < givingData.data.length; i++) {
                    totalAmount += givingData.data[i].totalAmount;
                }
                //This creates a tile for the return message
                var attachment1 = new builder.ThumbnailCard()
                    .title('Your Company\'s Volunteering Metrics:')
                    .text('<b>Total Hours:</b> ' + totalHours + '<br/>' +
                        '<b>Total Volunteers:</b> ' + totalParticipants + '<br/>')
                    .toAttachment()
                var attachment2 = new builder.ThumbnailCard()
                    .title('Your Company\'s Giving Metrics:')
                    .text('<b>Total Donations:</b> ' + totalAmount.toFixed(2) + '<br/>' +
                        '<b>Number of Donors:</b> ' + givingData.totalUniqueDonors + '<br/>')
                    .toAttachment()
                //This builds the message to send with the attachments
                var msg = new builder.Message(session)
                    .summary('Your Company\'s Donation and Volunteering Metrics')
                    .attachmentLayout('list') // carousel
                    .attachments([
                        attachment1,
                        attachment2
                    ]);
                session.send(msg);
            });
        } else if (text === 'now') {
            var promise1 = getUserDataWithPromise('https://api.yourcauseuat.com/v1/metrics/engagementelements');
            var promise2 = getUserDataWithPromise('https://api.yourcauseuat.com/v1/metrics/give/campaigns');
            Promise.all([promise1, promise2]).then(function (response) {
                var engagementElements = JSON.parse(response[0]);
                var giveCampaigns = JSON.parse(response[1]);

                var str = '';
                for (var i = 0; i < engagementElements.data.length; i++) {
                    str = str + '<b>Name:</b>' + ' ' + engagementElements.data[i].name + '<br/>';
                    str = str + '<b>Total Donors:</b>' + ' ' + engagementElements.data[i].totalDonors + '<br/>';
                    str = str + '<b>Total Amount:</b>' + ' ' + engagementElements.data[i].totalAmount + '<br/>';
                    //This is attaching the link that the user can use to go to more engagement
                    var attach = 'https://yc.yourcauseuat.com/home#/engagement/' + engagementElements.data[i].engagementElementId;
                    str += '<a href="' + attach + '">View the Campaign Here</a>';
                    str = str + '<br/>';
                }
                var attachment3 = new builder.ThumbnailCard()
                    .title('Help Now! These nonprofits need our help:')
                    .text(str)
                    .toAttachment()

                var str2 = '';
                for (var j = 0; j < giveCampaigns.data.length; j++) {
                    str2 = str2 + '<b>Charity Name:</b>' + ' ' + giveCampaigns.data[j].campaignName + '<br/>';
                    str2 = str2 + '<b>Total Donors:</b>' + ' ' + giveCampaigns.data[j].totalDonors + '<br/>';
                    str2 = str2 + '<b>Total Amount:</b>' + ' ' + giveCampaigns.data[j].totalAmount + '<br/>';
                    str2 = str2 + '<br/>';
                }
                str2 += '<a href="https://yc.yourcauseuat.com/home#/give/mygiving">Click here to Donate</a>';
                var attachment4 = new builder.ThumbnailCard()
                    .title('Donate to our company-wide giving campaigns:')
                    .text(str2)
                    .toAttachment()

                var msg = new builder.Message(session)
                    .summary('Active Now')
                    .attachmentLayout('list') // carousel
                    .attachments([
                        attachment3,
                        attachment4
                    ]);
                session.send(msg);
            });
        } else if (text === 'urgent') {
            getUserDataWithPromise('https://api.yourcauseuat.com/v1/metrics/engagementelements').then(function (result) {
                var ret = JSON.parse(result);
                var str = '';
                for (var i = 0; i < ret.data.length; i++) {
                    str = str + '<b>Name:</b>' + ' ' + ret.data[i].name + '<br/>';
                    str = str + '<b>Total Donors:</b>' + ' ' + ret.data[i].totalDonors + '<br/>';
                    str = str + '<b>Total Amount:</b>' + ' ' + ret.data[i].totalAmount + '<br/>';
                    var bone = 'https://yc.yourcauseuat.com/home#/engagement/' + ret.data[i].engagementElementId;
                    str += '<a href="' + bone +  '">View the Campaign Here</a>';
                    str = str + '<br/>';
                }
                var attachment = new builder.ThumbnailCard()
                    .title('Help Now! These nonprofits need our help:')
                    .text(str)
                    .toAttachment()
                var msg = new builder.Message(session)
                    .summary('Engagement Elements')
                    .attachmentLayout('list') // carousel
                    .attachments([
                        attachment
                    ]);
                session.send(msg);
            });
        } else if (text === 'active') {
            getUserDataWithPromise('https://api.yourcauseuat.com/v1/metrics/give/campaigns').then(function (result) {
                var ret = JSON.parse(result);
                var str = '';
                for (var i = 0; i < ret.data.length; i++) {
                    str = str + '<b>Charity Name:</b>' + ' ' + ret.data[i].campaignName + '<br/>';
                    str = str + '<b>Total Donors:</b>' + ' ' + ret.data[i].totalDonors + '<br/>';
                    str = str + '<b>Total Amount:</b>' + ' ' + ret.data[i].totalAmount + '<br/>';
                    str = str + '<br/>';
                }
                str += '<a href="https://yc.yourcauseuat.com/home#/give/mygiving">Click here to Donate</a>';
                var attachment = new builder.ThumbnailCard()
                    .title('Campaigns:')
                    .text(str)
                    .toAttachment()
                var msg = new builder.Message(session)
                    .summary('Campaigns')
                    .attachmentLayout('list') // carousel
                    .attachments([
                        attachment
                    ]);
                session.send(msg);
            });
        } else if (text === 'top charities') {
            getUserDataWithPromise('https://api.yourcauseuat.com/v1/metrics/charities').then(function (result) {
                var ret = JSON.parse(result);
                var str = '';
                for (var i = 0; i < ret.give.length; i++) {
                    str = str + '<b>Charity Name:</b>' + ' ' + ret.give[i].charityName + '<br/>';
                    str = str + '<b>Total Donors:</b>' + ' ' + ret.give[i].totalDonors + '<br/>';
                    str = str + '<b>Total Transaction:</b>' + ' ' + ret.give[i].totalTransactions + '<br/>';
                    str = str + '<b>Total Amount:</b>' + ' ' + ret.give[i].totalAmount + '<br/>';
                    str = str + '<br/>';
                }
                for (var i = 0; i < ret.volunteer.length; i++) {
                    str = str + '<b>Charity Name:</b>' + ' ' + ret.volunteer[i].charityName + '<br/>';
                    str = str + '<b>Total Opportunities:</b>' + ' ' + ret.volunteer[i].totalOpportunities + '<br/>';
                    str = str + '<b>Total Hours:</b>' + ' ' + ret.volunteer[i].totalHours + '<br/>';
                    str = str + '<b>Total Participants:</b>' + ' ' + ret.volunteer[i].totalParticipants + '<br/>';
                    str = str + '<br/>';
                }
                str += '<a href="https://yc.yourcauseuat.com/home#/newvolunteer">Click here to Volunteer</a>';
                var attachment = new builder.ThumbnailCard()
                    .title('Here are the top charities your Company helped this quarter!')
                    .text(str)
                    .toAttachment()
                var msg = new builder.Message(session)
                    .summary('Charities')
                    .attachmentLayout('list') // carousel
                    .attachments([
                        attachment
                    ]);
                session.send(msg);
            });
        } else if (text === 'help') {
            var attachment = new builder.ThumbnailCard()
                .title('Commands:')
                .text('<b> me:</b> Your Donation and Volunteering Metrics<br/>' +
                    '<b> company:</b> Company-wide Donation and Volunteering Metrics<br/>' +
                    '<b> urgent:</b> Engagement Elements- Urgent Giving Campaigns<br/>' +
                    '<b> active:</b> Company-Wide Active Giving Campaigns<br/>' +
                    '<b> now:</b> Engagement Elements and Company-Wide Active Giving Campaigns<br/>' +
                    '<b> top charities:</b> Top Charities by Volunteering and Donations Across the Company<br/>' +
                    '<b> help:</b> Explains All Commands<br/>')
                .toAttachment()
            var msg = new builder.Message(session)
                .summary('Grant applications')
                .attachmentLayout('list') // carousel
                .attachments([
                    attachment
                ]);
            session.send(msg);
        }
        
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
                    return result.userPrincipalName;
                    // session.endDialog('%s', JSON.stringify(result));
                }
            }
        );
    }

    //This is the function that calls the api to get the data
    //It hits the url that is passed to it
    //Returns a promise that will resolve to the API response text
    function getUserDataWithPromise(url) {
        var XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
        var xhr = new XMLHttpRequest();
        return new Promise(function (resolve, reject) {
            //Waits for the request to get to the state of a response and then resolves the response to the api response text
            xhr.onreadystatechange = function () {
                if (xhr.readyState == 4) {
                    if (xhr.status >= 300) {
                        reject("Error, status code = " + xhr.status)
                    } else {
                        resolve(xhr.responseText);
                    }
                }
            }
            xhr.open('get', url, true);
            xhr.setRequestHeader("x-bb-Key", "84704ed0-a429-4516-8a9d-fccab0bb49aa");
            xhr.send();
        });
    }

}
