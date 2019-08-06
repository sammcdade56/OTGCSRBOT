---
products:
- office-365
- Microsoft Teams
languages:
- javascript
- NodeJs
title: CSRobot
description: Microsoft Teams chatbot 
extensions:
  Created on: 08/05/2019
---
# CSRobot Setup

The bot is not published on the Microsoft App Marketplace; instead, you need to setup a bot either for a specific Team or for your account. To do this setup for a team, you must have a Microsoft Teams account with administrative permissions to add a bot to a channel. To do the setup for your account, you may need the same permissions, but we are not sure.

Follow the development set up at https://docs.microsoft.com/en-us/microsoftteams/platform/get-started/get-started-nodejs-app-studio until it asks you to run ‘gulp’ in the terminal (immediately after installing App Studio). Additionally, when you run the command “ngrok http 3333 -host-header=localhost:3333,” take note of the url. You will need this for connecting to the code.

1. Instead of importing an existing app, select create a new app.

2. Set the Short Name for the app as CSRobot

3. Click generate - you should get an App ID that looks like: 2322041b-72bf-459d-b107-f4f335bc35bd

4. You must fill out all the marked information on the remaining part of the App Details page. While you don’t have to match what we did, we set the Package Name to com.csrobot.bot, the version to 1.0.0, the Short description to “CSR reporting for a company's employees” the Long descritption to “Allows employees and companies to view CSR data easily, and enables reports.” The rest of the fields would be about YourCause.

5. Skip the tabs page under capabilities

6. Under the Bots page, add a new Bot. On the new bot tab, name the bot CSRobot and under scope select personal, team, and group chat.

7. Click Generate New Password, and make a note of the password in the same text file you noted your Bot app ID in. This is the only time your password will be shown, so be sure to do this now.

8. Update the Bot endpoint address to https://yourteamsapp.ngrok.io/api/messages, where yourteamsapp.ngrok.io should be replaced by the URL that you used above when hosting your app.

9. Follow the guide’s setup for messaging extensions (https://docs.microsoft.com/en-us/microsoftteams/platform/get-started/get-started-nodejs-app-studio)

10. Go to test and distribute, select install, and make sure the toggle for add to you is on yes, and if you want you can also add it to a team or specific chat. If you add it to a team, when you click install it will ask you which channels in the team to add the bot to. Select the ones you desire to add it to.

11. Now, to setup the bot, go to the code file and follow the guides instructions for updating the app for ngrok. For the app to be connected, you need to specifically update the BASE_URI to point to the url saved earlier, the MICROSOFT_APP_ID to match the app ID (note: this is different from the bot ID and should be found on the app details tab), and the MICROSOFT_APP_PASSWORD which you saved earlier.

12. Also change the instance of “ngrok.io” in the manifest.json and messaging-extension.js are pointing to the correct location.

From here, you can run npm start in the directory that your code is in, and the bot will be online.


## How to debug locally

1. Run ./ngrok http 3333 using the folloing command: ```-host-header=localhost:3333 ```
2. Replace the ngrok url in the teams config file.
3. Update the ngrok url in the connector config.
4. "Install" to the team.
5. Fix the oauth token.
6. Start debugging.
7. Must restart Teams if you restart the service to get link preview to work.

## Features
#### Completed Features

* Social Good statistics on-demand
  - Personal volunteering and giving totals
  - Company volunteering and giving totals
  - Top charities by volunteering and giving across the company
  - Engagement Elements- urgent giving campaigns
  - Company Give Campaigns

#### To-do List

* Implement notifications for each of the on-demand commands
* Log hours or donations via text
* Implement a recomendations engine based on location and stated interest
* Implement insights, tracking how the bot is being used
* Implement a competitions feature

# Official documentation

More information on =how to get started with Microsoft Teams development can be found in [Get started on the Microsoft Teams platform with Node.js and App Studio](https://docs.microsoft.com/en-us/microsoftteams/platform/get-started/get-started-nodejs-app-studio).

# File Conventions
  * ```src``` folder: Contains source code and images pertaining to the chatbot.
  * ```src/static/images``` folder: Folder containing images to be used in the bot's thumbnail.
  * ```src/static/scripts``` folder: Contains the initializing script, setting parameters in the bot to match the user's preferences within Teams.
  * ```src/static/styles``` folder: Contains the CSS files for the chatbot.
  * ```src/views``` folder: Contains a separate ```.pug``` file for each of the chatbot's tabs.
  * ```app.js``` file: Initializes the chatbot object, along with it's tabs and messaging extensions.
  * ```bot.js``` file: Includes the logic for each cmmand the chatbot supports.
  * ```tabs.js``` file: Includes the logic which renders each tab.

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

### Visual Studio Code extensions

The easiest way to deploy to Azure is to use Visual Studio Code with Azure extensions. There are many extensions for Azure - you can get all of them at once by installing the [Node Pack for Azure](https://marketplace.visualstudio.com/items?itemName=ms-vscode.vscode-node-azure-pack) or you can install just the [Azure App Service](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azureappservice) extension.

### Creating a new Node.js web app

Once you've installed the extensions, you'll see a new Azure icon on the left in Visual Studio Code. Click on the + icon to create a new web app. Once you've created your web app:

1. Add the following Application Settings (environment variables):

   ```
   MICROSOFT_APP_ID=<YOUR BOT'S APP ID>
   MICROSOFT_APP_PASSWORD=<YOUR BOT'S APP PASSWORD>
   WEBSITE_NODE_DEFAULT_VERSION=8.9.4
   ```

1. Configure the Deployment Source for your app (either your local copy of this repository or one you've forked on GitHub).
1. Deploy your web app. Visual Studio Code will tell you when you are done.

### Deploying to Azure for Node.js on Windows

Since this repo was optimized for Azure App Service, which runs on Linux, the `.deployment` file references `bash deploy.sh`. There's also a `deploy.cmd` if you want to deploy to Azure running Node.js on Windows. If you do, change `.deployment` to this instead:

```
[config]
command = deploy.cmd
```
