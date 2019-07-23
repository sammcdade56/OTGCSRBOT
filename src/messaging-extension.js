'use strict';

module.exports.setup = function() {
    var builder = require('botbuilder');
    var teamsBuilder = require('botbuilder-teams');
    var bot = require('./bot');
    const https = require('https');

    const subKey = '035f143314da4c2cb6b81542f30639c7';
    const authToken = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IjREVjZzVkxIM0FtU1JTbUZqMk04Wm5wWHU3WSJ9.eyJuYW1laWQiOiJhNjNkMjhjYy1hZGI4LTRlZGMtYmIzMS04ZTljZTBmZGE0NTUiLCJ0ZW5hbnRpZCI6IjhlNGYxYjU2LWMyZmUtNDE4OS1hN2VkLTg1NWQ3ODgxY2ExNiIsImFwcGxpY2F0aW9uaWQiOiJhMDU2Y2E2Yi1hM2E4LTRhYzctYjMyNS05OTc2NjYzMDZlNTIiLCJlbnZpcm9ubWVudGlkIjoidC16ZWFjeEx2cnprMlh4NXlzUWhQaWRBIiwiZW52aXJvbm1lbnRuYW1lIjoiQ2h1cmNoIERlbW8gRW52aXJvbm1lbnQgMSIsImxlZ2FsZW50aXR5aWQiOiJ0LU1ZS3JnTFRUdjBhUjFMakswbERhOUEiLCJsZWdhbGVudGl0eW5hbWUiOiJDaHVyY2ggRGVtbyIsImlzcyI6Imh0dHBzOi8vb2F1dGgyLnNreS5ibGFja2JhdWQuY29tLyIsImF1ZCI6IlJFeCIsImV4cCI6MTU2Mzg5NzYyMSwibmJmIjoxNTYzODk0MDIxfQ.Jn1gSI7TTyYK5FzbDJ3ZRxTVuGoa2_oze80DbipP6QpN3yPbaVyhWZlNFu7ZFO8pxRa_IDOpmvLUAWz2r2ivh6ozIwPJxnIV1pMrLQ2JC1KU0ElGaRT5IdyAR9EG463P-Ktv06_tMd0f_b_uvttGVC6SCdtoQqIOk7RYnmRPq5daVBskUCckf0hWetWVxUe7H0Yje70fBPe35r7zt-ertPfpFV33GId8AItoqjWgb8IW0zDjmpe58-c93E5het6ET5caF4UGOO52ck63JIG_Ttnf1U4IBq7P4XmwWx_5EGFp51fQPk9BPp7LfGxR7G5s6mcD07ed1P_zlmpi2C0V7w';
    const host = 'api.sky.blackbaud.com';
    const envId = 't-zeacxLvrzk2Xx5ysQhPidA'; // TODO get from authToken

    console.log('CES messaging-extension setup');

    bot.connector.onQuery('getRandomText', function(event, query, callback) {
      console.log('CES messaging-extension getRandomText');

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
          // console.log('CES resp: ' + resp);
          let data = '';

          // A chunk of data has been recieved.
          resp.on('data', (chunk) => {
            data += chunk;
          });

          // The whole response has been received. Print out the result.
          resp.on('end', () => {
            var dataObj = JSON.parse(data);
            // console.log('Result: ' + data);

            var constits = [];

            // https://developer.blackbaud.com/skyapi/apis/constituent/entities#SearchResult
            // id, address, deceased, email, fundraiser_status, inactive, lookup_id, name
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
                // console.log('Constits ' + i + constits.length);
                constits[i].thumbnailUrl = thumbnailUrl;
                waitingThumbnails--;

                if (waitingThumbnails === 0) {
                  completeSearch(constits, callback);
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

    function completeSearch(constits, callback) {

      // Build the data to send
      var attachments = [];

      for (var i = 0; i < constits.length; i++) {
        attachments.push(
          new builder.ThumbnailCard()
          .title(constits[i].title)
          .text(constits[i].text)
          .images([new builder.CardImage().url(constits[i].thumbnailUrl)])
          .tap({
            type: 'openUrl',
            title: 'Open constituent in RENXT',
            value: `https://renxt.blackbaud.com/constituents/${constits[i].id}?envid=${envId}`
          })
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
          // console.log('Result : ' + constits[i].id + data2);
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
