'use strict';

module.exports.setup = function() {
    var builder = require('botbuilder');
    var teamsBuilder = require('botbuilder-teams');
    var bot = require('./bot');
    const https = require('https');

    const subKey = '035f143314da4c2cb6b81542f30639c7';
    const authToken = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IjREVjZzVkxIM0FtU1JTbUZqMk04Wm5wWHU3WSJ9.eyJuYW1laWQiOiJhNjNkMjhjYy1hZGI4LTRlZGMtYmIzMS04ZTljZTBmZGE0NTUiLCJ0ZW5hbnRpZCI6IjhlNGYxYjU2LWMyZmUtNDE4OS1hN2VkLTg1NWQ3ODgxY2ExNiIsImFwcGxpY2F0aW9uaWQiOiJhMDU2Y2E2Yi1hM2E4LTRhYzctYjMyNS05OTc2NjYzMDZlNTIiLCJlbnZpcm9ubWVudGlkIjoidC16ZWFjeEx2cnprMlh4NXlzUWhQaWRBIiwiZW52aXJvbm1lbnRuYW1lIjoiQ2h1cmNoIERlbW8gRW52aXJvbm1lbnQgMSIsImxlZ2FsZW50aXR5aWQiOiJ0LU1ZS3JnTFRUdjBhUjFMakswbERhOUEiLCJsZWdhbGVudGl0eW5hbWUiOiJDaHVyY2ggRGVtbyIsImlzcyI6Imh0dHBzOi8vb2F1dGgyLnNreS5ibGFja2JhdWQuY29tLyIsImF1ZCI6IlJFeCIsImV4cCI6MTU2Mzg1Mjg5MSwibmJmIjoxNTYzODQ5MjkxfQ.OVzGJrzqccqDd6ENH56d_TF_ED0s-g_EMphK78kIqjo93Ht9QReSrmJ09vO1JneF4pjhE6Opt9b68tbzyCSbnr7-s13mWOkNNw86_Q8I7GHL3-j-Q9nE64E4Z5-Jg19LPWIN1rNMmfG03sXQrWnCGA9WxTWPTGJ5kzVWZZzhvAOk77fHq5LzsrYhleV1Hb3aPT8BUXm-DftPA50AfrBiYSxtcN35k3QwZJmtsd_ka-jN2zI772gdOUa-o4wK2JokUay1ygZtc_L2gU_COcUch7VTs9etZU-kdGjKXWsaYX_NzEhJoehK34WNqFF8dAEXZ7WxuvU_VpX2S8V8HrOWgA';
    const host = 'api.sky.blackbaud.com';

    console.log('CES messaging-extension setup');

    bot.connector.onQuery('getRandomText', function(event, query, callback) {
      console.log('CES messaging-extension getRandomText');

        var faker = require('faker');

        // If the user supplied a title via the cardTitle parameter then use it or use a fake title
        var searchText = query.parameters && query.parameters[0].name === 'cardTitle'
            ? query.parameters[0].value
            : ''; // faker.lorem.sentence();

        var options = {
          headers: {
            'Bb-Api-Subscription-Key': subKey,
            'Authorization': `Bearer ${authToken}`
          },
          method: 'GET',
          protocol: 'https:',
          defaultPort: 443,
          host: host,
          path: '/constituent/v1/constituents/search?search_text=' + encodeURIComponent(searchText)
        };
        const req = https.request(options, (resp) => {
          console.log('CES resp: ' + resp);
          let data = '';

          // A chunk of data has been recieved.
          resp.on('data', (chunk) => {
            data += chunk;
          });

          // The whole response has been received. Print out the result.
          resp.on('end', () => {
            var dataObj = JSON.parse(data);
            console.log('Result: ' + data);

            let randomImageUrl = "https://loremflickr.com/200/200"; // Faker's random images uses lorempixel.com, which has been down a lot

            var constits = [];

            for (var i = 0; i < Math.min(dataObj.count, 5); i++) {
              constits.push({
                id: dataObj.value[i].id,
                title: dataObj.value[i].name,
                text: dataObj.value[i].email
              });
            }

            if (constits.length === 0) {

              // Build the response to be sent
              var response = teamsBuilder.ComposeExtensionResponse
                  .result('list')
                  .attachments([
                    new builder.ThumbnailCard()
                      .title('No results')
                      .toAttachment()
                  ])
                  .toResponse();

              // Send the response to teams
              callback(null, response, 200);

            }

            var waitingThumbnails = constits.length;
            for (var i = 0; i < constits.length; i++) {
              getProfilePic(constits, i, (constits, i, thumbnailUrl) => {
                console.log('Constits ' + i + constits.length);
                constits[i].thumbnailUrl = thumbnailUrl;
                waitingThumbnails--;

                if (waitingThumbnails === 0) {

                  // Build the data to send
                  var attachments = [];

                  for (var i = 0; i < constits.length; i++) {
                    attachments.push(
                      new builder.ThumbnailCard()
                      .title(constits[i].title)
                      .text(constits[i].text)
                      .images([new builder.CardImage().url(constits[i].thumbnailUrl)])
                      .toAttachment());
                  }

                  // Build the response to be sent
                  var response = teamsBuilder.ComposeExtensionResponse
                      .result('list')
                      .attachments(attachments)
                      .toResponse();

                  // Send the response to teams
                  callback(null, response, 200);

                }

              });
            }

          });

        });

        req.on("error", (err) => {
          console.log("Error: " + err.message);
        });
        req.end();
    });

    function getProfilePic(constits, i, callback) {

      var options2 = {
        headers: {
          'Bb-Api-Subscription-Key': subKey,
          'Authorization': `Bearer ${authToken}`
        },
        method: 'GET',
        protocol: 'https:',
        defaultPort: 443,
        host: host,
        path: `/constituent/v1/constituents/${constits[i].id}/profilepicture`
      };

      https.request(options2, (resp2) => {
        // console.log('CES resp2: ' + JSON.stringify(resp2));
        let data2 = '';

        // A chunk of data has been recieved.
        resp2.on('data', (chunk) => {
          data2 += chunk;
        });

        // The whole response has been received. Print out the result.
        resp2.on('end', () => {
          var dataObj = JSON.parse(data2);
          console.log('Result : ' + constits[i].id + data2);
          var thumbnail = dataObj.thumbnail_url || 'https://upload.wikimedia.org/wikipedia/commons/8/89/Portrait_Placeholder.png';
          callback(constits, i, thumbnail);
        });
      })
      .on("error", (err) => {
        console.log("Error: " + err.message);
      })
      .end();
    }
};
