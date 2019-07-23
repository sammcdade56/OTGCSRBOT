'use strict';

module.exports.setup = function () {
  var builder = require('botbuilder');
  var teamsBuilder = require('botbuilder-teams');
  var bot = require('./bot');
  const https = require('https');

  const subKey = '035f143314da4c2cb6b81542f30639c7';
  const authToken = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IjREVjZzVkxIM0FtU1JTbUZqMk04Wm5wWHU3WSJ9.eyJuYW1laWQiOiJhNjNkMjhjYy1hZGI4LTRlZGMtYmIzMS04ZTljZTBmZGE0NTUiLCJ0ZW5hbnRpZCI6IjhlNGYxYjU2LWMyZmUtNDE4OS1hN2VkLTg1NWQ3ODgxY2ExNiIsImFwcGxpY2F0aW9uaWQiOiJhMDU2Y2E2Yi1hM2E4LTRhYzctYjMyNS05OTc2NjYzMDZlNTIiLCJlbnZpcm9ubWVudGlkIjoidC16ZWFjeEx2cnprMlh4NXlzUWhQaWRBIiwiZW52aXJvbm1lbnRuYW1lIjoiQ2h1cmNoIERlbW8gRW52aXJvbm1lbnQgMSIsImxlZ2FsZW50aXR5aWQiOiJ0LU1ZS3JnTFRUdjBhUjFMakswbERhOUEiLCJsZWdhbGVudGl0eW5hbWUiOiJDaHVyY2ggRGVtbyIsImlzcyI6Imh0dHBzOi8vb2F1dGgyLnNreS5ibGFja2JhdWQuY29tLyIsImF1ZCI6IlJFeCIsImV4cCI6MTU2MzkwNzUzMywibmJmIjoxNTYzOTAzOTMzfQ.RJrZgZXAsxjsrUabtDjihiRDRrKU-pxNgAKLeNvz6DLoj5dKMw7loVcRxLFPHLE1Xwfoa2zbGKQxm9xWUZ2aHyD9nOOZRIybelvVJyLVeUN7MuGdC3uMWG6IXzXAL2Gayi8PPhCuKHoFTALLMclo0I0U7TSUgSXBvJUUNayCCNnoFgtKg-je6NAo6o8kQoMz-B7o0qoWMj__9wqtAgeTOZNvSK-vAODo_5YMjWYW0awvzJQmGzWsQNFbSwKHoTlhzuB7hhDf-dO1h9-gCzIWpwEqGKTbMvMTJIlwj4oNkWOo_du1SSBaDAaWuP1B8VQm0aa07cOfucshGn33zdLsFg';
  const host = 'api.sky.blackbaud.com';
  const envId = 't-zeacxLvrzk2Xx5ysQhPidA'; // TODO get from authToken

  console.log('CES messaging-extension setup');

  bot.connector.onInvoke(function (event, callback) {
    console.log('CES messaging-extension addAction ' + JSON.stringify(event));
    try {
      if (event.name == 'composeExtension/fetchTask') {
        // No idea what to do here, see readme
          // bot.loadSession(event.address, (err, session) => {
          //     let verificationCode = event.value.state;
          //     // Get the user token using the verification code sent by MS Teams
          //     connector.getUserToken(session.message.address, connectionName, verificationCode, (err, result) => {
          //         session.send('Token ' + result.token);
          //         session.userData.activeSignIn = false;
                  callback(undefined, {}, 200);
          //     });
          // });
      } else {
          callback(undefined, {}, 200);
      }
    } catch (err) {
        console.log(err);
    }
  });

  bot.connector.onQuery('constituentSearch', function (event, query, callback) {
    console.log('CES messaging-extension constituentSearch');

    var searchText = query.parameters && query.parameters[0].name === 'searchText'
      ? query.parameters[0].value
      : '';

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
            name: dataObj.value[i].name,
            email: dataObj.value[i].email
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
        var waitingStatuses = constits.length;

        for (var i = 0; i < constits.length; i++) {
          getProfilePic(constits, i, (constits, i, thumbnailUrl) => {
            constits[i].thumbnailUrl = thumbnailUrl;
            waitingThumbnails--;

            if (constitSearchIsComplete(waitingThumbnails, waitingStatuses)) {
              completeSearch(constits, callback);
            }
          });
        }

        for (var i = 0; i < constits.length; i++) {
          getProspectStatus(constits, i, (constits, i, status) => {
            constits[i].status = status;
            waitingStatuses--;

            if (constitSearchIsComplete(waitingThumbnails, waitingStatuses)) {
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

    function constitSearchIsComplete(waitingThumbnails, waitingStatuses) {
      return waitingThumbnails === 0 &&
        waitingStatuses === 0;
    }

    function completeSearch(constits, callback) {

      // Build the data to send
      var attachments = [];

      for (var i = 0; i < constits.length; i++) {
        attachments.push(
          new builder.ThumbnailCard()
            .title(constits[i].name)
            .subtitle('<a href="mailto:' + constits[i].email + '">' + constits[i].email + '</a>')
            .text('Prospect status: ' + constits[i].status + '<br/>line 2<br/> line 3')
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

    function getProspectStatus(constits, i, callback) {

      var options2 = {
        headers: {
          'Bb-Api-Subscription-Key': subKey,
          'Authorization': `Bearer ${authToken}`
        },
        method: 'GET',
        protocol: 'https:',
        defaultPort: 443,
        host: host,
        path: `/constituent/v1/constituents/${constits[i].id}/prospectstatus`
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
          // https://developer.blackbaud.com/skyapi/apis/constituent/entities#ProspectStatus
          var dataObj = JSON.parse(data2);
          console.log('Result : ' + constits[i].id + data2);
          callback(constits, i, dataObj.status || 'N/A');
        });
      })
        .on("error", (err) => {
          console.log("Error: " + err.message);
        })
        .end();
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

  });

};
