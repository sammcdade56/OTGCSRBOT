## Teams developer complaints

* Documentation appears out of date with UI - screenshots and instructions don't match, lots of new fields and not sure what they do
* No good way for multiple to work on the same app that I could find
* I have to manually go into Teams and delete, import, install my app if I update the manifest
* No indication that the app has been updated but not installed in App Studio
* Can't copy text from error message in App Studio (took screenshot)
* `/manifestVersion must be equal to...`
* Can't test tabs inside app studio
* Link stuff just doesn't work sometimes, no idea why. I can fix it by restarting teams, which is stupid.
* `<br/>` line breaks work for card thumbnail in chatbot message but not as page preview
* Can I have a different card shown when clicking results from a list?
* Excluding `.text` from the following in the app link response makes it do nothing, and not show a console error
```
var response = teamsBuilder.ComposeExtensionResponse
    .result('list')
    .attachments([
      new builder.ThumbnailCard()
        .title('Constituent preview coming soon!')
        .text('coming soon')
        .toAttachment()
    ])
    .toResponse();
```

## todo

[Authentication](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/messaging-extensions/search-extensions#authentication)

### constituent search action

#### Done

* Search RENXT's constituent search
* Display top 5 results
* Message if empty result
* Show profile pic
* Show email address
* Show prospect status

#### Not done

* Can I show a different thing in the results list and on the card that is shown?
* latest gift
* lifetime giving
* last action (list actions, latest date)
* authentication
  * can we avoid the timeout?
* why is it showing up twice when you click what can I do?

### bot

#### create an action
* docs are wrong/unhelpful
  * `addTodo` on https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/messaging-extensions/create-extensions
    * When this exact json is loaded, the UI in app studio does not match - parameters are not there
  * No examples of how to hook up `connector.onInvoke`
  * No information about what the invoke handler is supposed to return
* I have no idea what to do with this

### link

#### Done - Preview `renxt.blackbaud.com/constituent/{id}` when pasted into chat
* see `exampleMessageHandler` on https://docs.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema
* docs are wrong/unhelpful
  * https://docs.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema lists `messageHandlers` which is not in the schema
  * Seem to be correct: https://developer.microsoft.com/en-us/yammer/blogs/add-rich-previews-to-messages-using-link-unfurling/
  * Ran into issue with `.text` documented above

### events

get notified when someone adds an action to my prospect

## how to debug locally

Run ./ngrok http 3333 -host-header=localhost:3333
Replace the ngrok url in the teams config
"Install" to the team
Fix the oauth token
Start debugging

## how to upload the app

prereq:
`npm i -g gulp

gulp
Open Teams' App studio
Manifest editor
Delete the one that's already there?
Import an existing app
Find `./manifest/blackbaudteamsapp.zip`

---
topic: sample
products:
- office-365
languages:
- javascript
extensions:
  contentType: samples
  createdDate: 11/3/2017 12:53:17 PM
---
# Official documentation

More information for this sample - and for how to get started with Microsoft Teams development in general - is found in [Get started on the Microsoft Teams platform with Node.js and App Studio](https://docs.microsoft.com/en-us/microsoftteams/platform/get-started/get-started-nodejs-app-studio).

# Using this sample locally

This sample can be run locally using `ngrok` as described in the [official documentation](https://docs.microsoft.com/en-us/microsoftteams/platform/get-started/get-started-nodejs-app-studio), but you'll need to set up some environment variables. There are many ways to do this, but the easiest, if you are using Visual Studio Code, is to add a [launch configuration](https://code.visualstudio.com/Docs/editor/debugging#_launch-configurations):

```json
[...]
        {
            "type": "node",
            "request": "launch",
            "name": "Launch - Teams Debug",
            "program": "${workspaceRoot}/src/app.js",
            "cwd": "${workspaceFolder}/src",
            "env": {
                "BASE_URI": "https://########.ngrok.io",
                "MICROSOFT_APP_ID": "00000000-0000-0000-0000-000000000000",
                "MICROSOFT_APP_PASSWORD": "yourBotAppPassword",
                "NODE_DEBUG": "botbuilder",
                "SUPPRESS_NO_CONFIG_WARNING": "y",
                "NODE_CONFIG_DIR": "../config"
            }
[...]
```

Where:

* `########` matches your actual ngrok URL
* `MICROSOFT_APP_ID` and `MICROSOFT_APP_PASSWORD` is the ID and password, respectively, for your bot
* `NODE_DEBUG` will show you what's happening in your bot in the Visual Studio Code debug console
* `NODE_CONFIG_DIR` points to the directory at the root of the repository (by default, when the app is run locally, it looks for it in the `src` folder)

# Deploying to Azure App Service

## Visual Studio Code extensions

The easiest way to deploy to Azure is to use Visual Studio Code with Azure extensions. There are many extensions for Azure - you can get all of them at once by installing the [Node Pack for Azure](https://marketplace.visualstudio.com/items?itemName=ms-vscode.vscode-node-azure-pack) or you can install just the [Azure App Service](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azureappservice) extension.

## Creating a new Node.js web app

Once you've installed the extensions, you'll see a new Azure icon on the left in Visual Studio Code. Click on the + icon to create a new web app. Once you've created your web app:

1. Add the following Application Settings (environment variables):

   ```
   MICROSOFT_APP_ID=<YOUR BOT'S APP ID>
   MICROSOFT_APP_PASSWORD=<YOUR BOT'S APP PASSWORD>
   WEBSITE_NODE_DEFAULT_VERSION=8.9.4
   ```

1. Configure the Deployment Source for your app (either your local copy of this repository or one you've forked on GitHub).
1. Deploy your web app. Visual Studio Code will tell you when you are done.

## Deploying to Azure for Node.js on Windows

Since this repo was optimized for Azure App Service, which runs on Linux, the `.deployment` file references `bash deploy.sh`. There's also a `deploy.cmd` if you want to deploy to Azure running Node.js on Windows. If you do, change `.deployment` to this instead:

```
[config]
command = deploy.cmd
```

# Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
=======
# Official documentation

More information for this sample - and for how to get started with Microsoft Teams development in general - is found in [Get started on the Microsoft Teams platform with Node.js and App Studio](https://docs.microsoft.com/en-us/microsoftteams/platform/get-started/get-started-nodejs-app-studio).

# Using this sample locally

This sample can be run locally using `ngrok` as described in the [official documentation](https://docs.microsoft.com/microsoftteams/platform/get-started/get-started-nodejs-app-studio), but you'll need to set up some environment variables. There are many ways to do this, but the easiest, if you are using Visual Studio Code, is to add a [launch configuration](https://code.visualstudio.com/Docs/editor/debugging#_launch-configurations):

```json
[...]
        {
            "type": "node",
            "request": "launch",
            "name": "Launch - Teams Debug",
            "program": "${workspaceRoot}/src/app.js",
            "cwd": "${workspaceFolder}/src",
            "env": {
                "BASE_URI": "https://########.ngrok.io",
                "MICROSOFT_APP_ID": "00000000-0000-0000-0000-000000000000",
                "MICROSOFT_APP_PASSWORD": "yourBotAppPassword",
                "NODE_DEBUG": "botbuilder",
                "SUPPRESS_NO_CONFIG_WARNING": "y",
                "NODE_CONFIG_DIR": "../config"
            }
[...]
```

Where:

* `########` matches your actual ngrok URL
* `MICROSOFT_APP_ID` and `MICROSOFT_APP_PASSWORD` is the ID and password, respectively, for your bot
* `NODE_DEBUG` will show you what's happening in your bot in the Visual Studio Code debug console
* `NODE_CONFIG_DIR` points to the directory at the root of the repository (by default, when the app is run locally, it looks for it in the `src` folder)

# Deploying to Azure App Service

## Visual Studio Code extensions

The easiest way to deploy to Azure is to use Visual Studio Code with Azure extensions. There are many extensions for Azure - you can get all of them at once by installing the [Node Pack for Azure](https://marketplace.visualstudio.com/items?itemName=ms-vscode.vscode-node-azure-pack) or you can install just the [Azure App Service](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azureappservice) extension.

## Creating a new Node.js web app

Once you've installed the extensions, you'll see a new Azure icon on the left in Visual Studio Code. Click on the + icon to create a new web app. Once you've created your web app:

1. Add the following Application Settings (environment variables):

   ```
   MICROSOFT_APP_ID=<YOUR BOT'S APP ID>
   MICROSOFT_APP_PASSWORD=<YOUR BOT'S APP PASSWORD>
   WEBSITE_NODE_DEFAULT_VERSION=8.9.4
   ```

1. Configure the Deployment Source for your app (either your local copy of this repository or one you've forked on GitHub).
1. Deploy your web app. Visual Studio Code will tell you when you are done.

## Deploying to Azure for Node.js on Windows

Since this repo was optimized for Azure App Service, which runs on Linux, the `.deployment` file references `bash deploy.sh`. There's also a `deploy.cmd` if you want to deploy to Azure running Node.js on Windows. If you do, change `.deployment` to this instead:

```
[config]
command = deploy.cmd
```

# Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

