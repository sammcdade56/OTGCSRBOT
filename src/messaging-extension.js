'use strict';

module.exports.setup = function () {
  var builder = require('botbuilder');
  var teamsBuilder = require('botbuilder-teams');
  var bot = require('./bot');
  const https = require('https');

  const subKey = '035f143314da4c2cb6b81542f30639c7';
  const authToken = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IjREVjZzVkxIM0FtU1JTbUZqMk04Wm5wWHU3WSJ9.eyJuYW1laWQiOiJhNjNkMjhjYy1hZGI4LTRlZGMtYmIzMS04ZTljZTBmZGE0NTUiLCJ0ZW5hbnRpZCI6IjhlNGYxYjU2LWMyZmUtNDE4OS1hN2VkLTg1NWQ3ODgxY2ExNiIsImFwcGxpY2F0aW9uaWQiOiJhMDU2Y2E2Yi1hM2E4LTRhYzctYjMyNS05OTc2NjYzMDZlNTIiLCJlbnZpcm9ubWVudGlkIjoidC16ZWFjeEx2cnprMlh4NXlzUWhQaWRBIiwiZW52aXJvbm1lbnRuYW1lIjoiQ2h1cmNoIERlbW8gRW52aXJvbm1lbnQgMSIsImxlZ2FsZW50aXR5aWQiOiJ0LU1ZS3JnTFRUdjBhUjFMakswbERhOUEiLCJsZWdhbGVudGl0eW5hbWUiOiJDaHVyY2ggRGVtbyIsImlzcyI6Imh0dHBzOi8vb2F1dGgyLnNreS5ibGFja2JhdWQuY29tLyIsImF1ZCI6IlJFeCIsImV4cCI6MTU2MzkyNDM2NiwibmJmIjoxNTYzOTIwNzY2fQ.KD7lQ5pxy2aUBf2riH7o8JFi1NwywAAH2J41R6nYyrOOGppNEHWKZsY_Vik0oXNvV8eKmNcGQZrJdhN3eUMro0VUQ9jN55ySyaxj02F25iBDdrQ1Nw3ehvamY6imQ5IcMFXx_JneSLChZkjRA4IGpojBdb00tvTcOks3SKtuo300Kj5FH1v26nPBqPMU8GsOP5Her5652v8Rti3Po13P7HH_KrgShqyklzydzx_nE88r-vY6HqwouTsGBF4j9H99gULQv0N11N1Z0P3x3oIQbYsVNDSYNvqt6WZ2VmWITEPfUdUFTa3Jp8Ly8SILIPtg8GesmF5OLUfIgxrUrnPo4w';
  const host = 'api.sky.blackbaud.com';
  const envId = 't-zeacxLvrzk2Xx5ysQhPidA'; // TODO get from authToken

  console.log('CES messaging-extension setup');

  bot.connector.onInvoke(function (event, callback) {
    console.log('CES messaging-extension onInvoke ' + JSON.stringify(event));
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

  bot.connector.onAppBasedLinkQuery(function (event, query, callback) {
    console.log('CES messaging-extension onAppBasedLinkQuery '); // + JSON.stringify(event) + ', query: ' + JSON.stringify(query));
    // query.url = 'https://renxt.blackbaud.com/constituents/280?tenantid=8e4f1b56-c2fe-4189-a7ed-855d7881ca16&svcid=chrch'

    var url = new URL(query.url);

    if (url.host === 'renxt.blackbaud.com') {
      if (url.pathname.substring(0, url.pathname.lastIndexOf('/') + 1) === '/constituents/') {
        var constitId = url.pathname.substring(url.pathname.lastIndexOf('/') + 1);

        var gotDetails = false;
        var gotProfilePic = false;
        var gotProspectStatus = false;

        var constituent = {
          id: constitId
        };

        getConstituentDetails(constitId, (details) => {
          constituent.name = details.name;
          constituent.email = details.email;
          gotDetails = true;

          if (gotDetails && gotProfilePic && gotProspectStatus) {
            completeResponse();
          }
        });

        getProfilePicForConstituent(constitId, (thumbnailUrl) => {
          constituent.thumbnailUrl = thumbnailUrl;
          gotProfilePic = true;

          if (gotDetails && gotProfilePic && gotProspectStatus) {
            completeResponse();
          }
        });

        getProspectStatusForConstituent(constitId, (status) => {
          constituent.status = status;
          gotProspectStatus = true;

          if (gotDetails && gotProfilePic && gotProspectStatus) {
            completeResponse();
          }
        });

      }

      function completeResponse() {
        var response = teamsBuilder.ComposeExtensionResponse
          .result('list')
          .attachments([
            getConstituentAttachment(constituent)
          ])
          .toResponse();

        callback(null, response, 200);
      }
    }

    // CES host: renxt.blackbaud.com
    // CES hostname: renxt.blackbaud.com
    // CES href: https://renxt.blackbaud.com/constituents/280?tenantid=8e4f1b56-c2fe-4189-a7ed-855d7881ca16&svcid=chrch
    // CES origin: https://renxt.blackbaud.com
    // CES pathname: /constituents/280
    // CES port:
    // CES protocol: https:
    // CES search: ?tenantid=8e4f1b56-c2fe-4189-a7ed-855d7881ca16&svcid=chrch
    // CES searchParams: tenantid=8e4f1b56-c2fe-4189-a7ed-855d7881ca16&svcid=chrch

  });

  bot.connector.onComposeExtensionFetchTask(function (event, request, callback) {
    console.log('CES messaging-extension onComposeExtensionFetchTask ');
  });

  bot.connector.onComposeExtensionSubmitAction(function (event, request, callback) {
    console.log('CES messaging-extension onComposeExtensionSubmitAction ');
  });

  bot.connector.onEvent(function (events, callback) {
    console.log('CES messaging-extension onEvent ' + JSON.stringify(events));
  });

  bot.connector.onO365ConnectorCardAction(function (event, query, callback) {
    console.log('CES messaging-extension onO365ConnectorCardAction ');
  });

  bot.connector.onQuerySettingsUrl(function (event, query, callback) {
    console.log('CES messaging-extension onQuerySettingsUrl ');
  });

  bot.connector.onSelectItem(function (event, query, callback) {
    console.log('CES messaging-extension onSelectItem ');
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
          getProfilePicForConstituents(constits, i, (constits, i, thumbnailUrl) => {
            constits[i].thumbnailUrl = thumbnailUrl;
            waitingThumbnails--;

            if (constitSearchIsComplete(waitingThumbnails, waitingStatuses)) {
              completeSearch(constits, callback);
            }
          });
        }

        for (var i = 0; i < constits.length; i++) {
          getProspectStatusForConstituents(constits, i, (constits, i, status) => {
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
        attachments.push(getConstituentAttachment(constits[i]));
      }

      // Build the response to be sent
      var response = teamsBuilder.ComposeExtensionResponse
        .result('list')
        .attachments(attachments)
        .toResponse();

      // Send the response to teams
      callback(null, response, 200);
    }

    function getProspectStatusForConstituents(constits, i, callback) {
      getProspectStatusForConstituent(constits[i].id, function(status) {
        callback(constits, i, status);
      });
    }

    function getProfilePicForConstituents(constits, i, callback) {
      getProfilePicForConstituent(constits[i].id, function(thumbnail) {
        callback(constits, i, thumbnail);
      });
    }
  });

  function getConstituentAttachment(constituent) {
    return new builder.ThumbnailCard()
      .title(constituent.name)
      .subtitle('<a href="mailto:' + constituent.email + '">' + constituent.email + '</a>')
      .text('Prospect status: ' + constituent.status + '<br/>line 2<br/>line 3')
      // .text('Prospect status: ' + constituent.status + '\nline 2\nline 3')
      .images([new builder.CardImage().url(constituent.thumbnailUrl)])
      .tap({
        type: 'openUrl',
        title: 'Open constituent in RENXT',
        value: `https://renxt.blackbaud.com/constituents/${constituent.id}?envid=${envId}`
      })
      .toAttachment();
  }

  function getConstituentDetails(constitId, callback) {

    var options = {
      headers: {
        'Bb-Api-Subscription-Key': subKey,
        'Authorization': `Bearer ${authToken}`
      },
      method: 'GET',
      protocol: 'https:',
      defaultPort: 443,
      host: host,
      path: `/constituent/v1/constituents/${constitId}`
    };

    https.request(options, (resp2) => {
      let data2 = '';

      // A chunk of data has been recieved.
      resp2.on('data', (chunk) => {
        data2 += chunk;
      });

      // The whole response has been received. Print out the result.
      resp2.on('end', () => {
        var dataObj = JSON.parse(data2);
        // https://developer.blackbaud.com/skyapi/apis/constituent/entities#Constituent
        var details = {
          email: dataObj.email,
          name: dataObj.name
        };
        callback(details);
      });
    })
    .on("error", (err) => {
      console.log("Error: " + err.message);
    })
    .end();
  }

  function getProfilePicForConstituent(constitId, callback) {

    var options = {
      headers: {
        'Bb-Api-Subscription-Key': subKey,
        'Authorization': `Bearer ${authToken}`
      },
      method: 'GET',
      protocol: 'https:',
      defaultPort: 443,
      host: host,
      path: `/constituent/v1/constituents/${constitId}/profilepicture`
    };

    https.request(options, (resp2) => {
      // console.log('CES resp2: ' + JSON.stringify(resp2));
      let data2 = '';

      // A chunk of data has been recieved.
      resp2.on('data', (chunk) => {
        data2 += chunk;
      });

      // The whole response has been received. Print out the result.
      resp2.on('end', () => {
        var dataObj = JSON.parse(data2);
        var thumbnail = dataObj.thumbnail_url || 'https://upload.wikimedia.org/wikipedia/commons/8/89/Portrait_Placeholder.png';
        callback(thumbnail);
      });
    })
    .on("error", (err) => {
      console.log("Error: " + err.message);
    })
    .end();
  }

  function getProspectStatusForConstituent(constitId, callback) {

    var options = {
      headers: {
        'Bb-Api-Subscription-Key': subKey,
        'Authorization': `Bearer ${authToken}`
      },
      method: 'GET',
      protocol: 'https:',
      defaultPort: 443,
      host: host,
      path: `/constituent/v1/constituents/${constitId}/prospectstatus`
    };

    https.request(options, (resp2) => {
      let data2 = '';

      // A chunk of data has been recieved.
      resp2.on('data', (chunk) => {
        data2 += chunk;
      });

      // The whole response has been received. Print out the result.
      resp2.on('end', () => {
        // https://developer.blackbaud.com/skyapi/apis/constituent/entities#ProspectStatus
        var dataObj = JSON.parse(data2);
        callback(dataObj.status || 'N/A');
      });
    })
    .on("error", (err) => {
      console.log("Error: " + err.message);
    })
    .end();
  }

};
