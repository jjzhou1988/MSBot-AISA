/*-----------------------------------------------------------------------------
A simple echo bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/
require('console');
var restify = require('restify');
var builder = require('botbuilder');
var botbuilder_azure = require("botbuilder-azure");
var teams = require("botbuilder-teams");
var rq = require('request-promise');


// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function() {
    console.log('%s listening to %s', server.name, server.url);
});

// Create chat connector for communicating with the Bot Framework Service
/*
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    openIdMetadata: process.env.BotOpenIdMetadata 
});
*/

//Create Teams Chat Bot
var connector = new teams.TeamsChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());

var bot = new builder.UniversalBot(connector);

/*----------------------------------------------------------------------------------------
 * Bot Storage: This is a great spot to register the private state storage for your bot. 
 * We provide adapters for Azure Table, CosmosDb, SQL Azure, or you can implement your own!
 * For samples and documentation, see: https://github.com/Microsoft/BotBuilder-Azure
 * ---------------------------------------------------------------------------------------- */

var tableName = 'botdata';
var azureTableClient = new botbuilder_azure.AzureTableClient(tableName, process.env['AzureWebJobsStorage']);
var tableStorage = new botbuilder_azure.AzureBotStorage({ gzipData: false }, azureTableClient);

// Create your bot with a function to receive messages from the user
//var bot = new builder.UniversalBot(connector);
bot.set('storage', tableStorage);
var stripBotAtMentions = new teams.StripBotAtMentions();
bot.use(stripBotAtMentions);

var count = 0;
var helpMsg = "Help message: <br>1. Please follow the hints and you will get what you want <br>2. Type quit to start over <br>3. Contact jinjiez@microsoft.com for the query";
//var helpMsg="1. Please follow the hints and you will get what you want <br>2. Type quit to start over <br>3. Contact jinjiez@microsoft.com and ashhu@microsoft.com for the query";
var dataServicePoint = 'https://cssadvisoryapiapp.azurewebsites.net';
var luisEndpoint = 'https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/2c8a694d-30b9-4203-a0f2-5d03815ccea4?subscription-key=c82d1549c7f145ceb7d832468201575d&verbose=true&timezoneOffset=0&q=';
var servicesType = ['teaminfo', 'techinfo','newfeature'];

// Authoring key, available in luis.ai under Account Settings
var LUIS_authoringKey = "c82d1549c7f145ceb7d832468201575d";
// ID of your LUIS app to which you want to add an utterance
var LUIS_appId = "2c8a694d-30b9-4203-a0f2-5d03815ccea4";
// The version number of your LUIS app
var LUIS_versionId = "0.1";

/*
var oneNoteImageUrl = '.\\Resources\\picto_one_note.jpg';
var wikiImageUrl = '.\\Resources\\picto_wiki.jpeg';
*/

var oneNoteImageUrl = 'https://rocketbotfy18.azurewebsites.net/picto_one_note.jpg';
var wikiImageUrl = 'https://rocketbotfy18.azurewebsites.net/WikipediaIcon.png';
var ser_type_querystr = {
    'teaminfo': 'teaminfo',
    'techinfo': 'allinfo'
};

var configAddUtterance = {
    LUIS_authoringKey: LUIS_authoringKey,
    LUIS_appId: LUIS_appId,
    LUIS_versionId: LUIS_versionId,
    uri: "https://westus.api.cognitive.microsoft.com/luis/api/v2.0/apps/{appId}/versions/{versionId}/examples".replace("{appId}", LUIS_appId).replace("{versionId}", LUIS_versionId)
};

function constructJsonUtterance(text, intent) {
    var jsonstr = [{
        "text": text,
        "intentName": intent,
        "entityLabels": []
    }];
    return jsonstr;
}

function generateHtmlTeamDetail(jsonContent) {

    var result = "";

    for (var item in jsonContent) {
        if (!(["Id", "TeamName", "TeamManager", "Keywords"].includes(item))) {
            console.log(item);
            result += (
                `<h2> ${item} </h2>` +
                "<ul>"
            );

            var entries = jsonContent[item].split(';');
            entries.forEach(element => {
                if (element != "") {
                    result += `<li> ${element} </li>`;
                }

            });
            result += "</ul>";
        }
    }
    return result;
}

function data_query(session, ser_type, ser_key) {
    console.log("enter data query func#");
    var uri = dataServicePoint + '/' + ser_type_querystr[ser_type] + '/' + encodeURIComponent(ser_key);
    console.log("uri : %s", uri);
    var options = {
        uri: uri,
        headers: {
            'User-Agent': 'Request-Promise'
        },
        simple: false,
        json: true // Automatically parses the JSON string in the response
    };

    rq(options)
        .then(function(repos) {

            if (ser_type == 'teaminfo') {
                console.log("TeamInfo : Responses got from data query : %s", JSON.stringify(repos));
                var records = repos;
                console.log('%d results matched', records.length);
                if (records.length >= 1) {
                    session.conversationData.teamInfoData[ser_key] = records;
                    session.conversationData.stage = 3;
                } else {
                    session.conversationData.teamInfoData[ser_key] = 'fail';
                    session.conversationData.stage = 4;
                }
                session.beginDialog(ser_type);
            } else if (ser_type == 'techinfo') {
                var records = repos;
                session.conversationData.techref = records;
                session.beginDialog(ser_type);
            }
        })
        .catch(function(err) {
            // Crawling failed...
        });
}


bot.dialog('/', [

    //check if it's a new session, if yes - welcome. Otherwise, go ahead
    function(session, args, next) {
        session.sendTyping();
        console.log('enter root dialogue water fall func#1');
        console.log('session.conversationData.stage : %d', session.conversationData.stage);
        if (!session.conversationData.stage) {
            session.conversationData.stage = 0;
            session.beginDialog('greetings');
        }
    },

    function(session) {
        console.log('enter root dialogue water fall func#2');
        var serv = session.conversationData.service;
        session.beginDialog(serv);
    }
]);

bot.dialog('greetings', [
    function(session, args, next) {
        console.log("session.message.text: %s", session.message.text);
        if (servicesType.indexOf(session.message.text) != -1) {
            next();
        } else {
            console.log('enter greetings water fall func#1');
            var title = "A.I.S.A is at your service~";
            var text = "Please choose the service you need. Type 'help' if any assistance is required.";
            var msg = new builder.Message(session);
            msg.attachments([
                new builder.ThumbnailCard(session)
                .title(title)
                .text(text)
                .buttons([
                    builder.CardAction.messageBack(session, '', 'Team Contacts')
                    .text('teaminfo'),
                    builder.CardAction.messageBack(session, '', 'New Features(comming soon)')
                    .text('newfeature'),
                ])
            ]);
            session.send(msg);
        }
    },
    function(session, results) {
        if (servicesType.indexOf(session.message.text) != -1) {
            console.log('enter greetings water fall func#2');
            session.conversationData.service = session.message.text;
            session.conversationData.stage = 1;
            session.conversationData.teamInfoData = {};
            session.conversationData.retry = 0;
            session.conversationData.currentKey;
            session.conversationData.techRef = [];
            session.conversationData.intent;
            session.conversationData.techDescription;
            session.endDialog();
        }
    }
]);

bot.dialog('help', [
        function(session) {
            console.log('enter help func');
            session.sendTyping();
            session.send(helpMsg).endDialog();
        }
    ])
    .triggerAction({
        matches: /^help$/i,
        onSelectAction: (session, args, next) => {
            // Add the help dialog to the dialog stack 
            // (override the default behavior of replacing the stack)
            session.beginDialog(args.action, args);
        }
    });

bot.dialog('quit', [
        function(session) {
            console.log('enter quit func');
            session.endConversation("Quit Conversation");
        }
    ])
    .triggerAction({
        matches: /^quit/i,
        onSelectAction: (session, args, next) => {
            // Add the help dialog to the dialog stack 
            // (override the default behavior of replacing the stack)
            session.beginDialog(args.action, args);
        }
    });



bot.dialog('teaminfo', [
    function(session, args, next) {
        session.sendTyping();
        console.log('enter teaminfo water fall func#1');
        console.log('session.conversationData.stage : %d', session.conversationData.stage);
        if (session.conversationData.stage == 1) {
            builder.Prompts.text(session, "Please tell me which team you want to contact?");
        } else if (session.conversationData.stage >= 2 && session.conversationData.stage <= 4) {
            next();
        } else {
            console.log("sth wrong !  session.conversationData.stage = %d ", session.conversationData.stage);
        }
    },
    function(session, results, next) {
        console.log("session.message.text : %s", session.message.text);
        console.log("results.response : %s", results.response);
        console.log('enter teaminfo water fall func#2');
        console.log('session.conversationData.stage : %d', session.conversationData.stage);

        if (session.conversationData.stage >= 3) {
            next();
        } else {
            var key = session.conversationData.currentKey = session.message.text.trim();

            if (session.conversationData.stage == 1) {
                /*key=results.response;*/
                session.conversationData.stage++;
            }

            console.log("still in teaminfo water fall func#2 key : %s", key);

            if (!(key in session.conversationData.teamInfoData)) {
                data_query(session, "teaminfo", key);
                session.endDialog();
            } else if (session.conversationData.teamInfoData[key] == 'fail') {
                session.conversationData.stage = 4;
                next();
            } else {
                session.conversationData.stage = 3;
                next();
            }
        }
    },
    function(session, results, next) {
        console.log('enter teaminfo water fall func#3');
        console.log('session.conversationData.stage : %d', session.conversationData.stage);
        if (session.conversationData.stage == 3) {
            var key = session.conversationData.currentKey;
            var records = session.conversationData.teamInfoData[key];

            var msg = new builder.Message(session);

            var cards = [];

            records.forEach(function(item, index) {
                var jsonContent = item;
                console.log('index : %d ; jsonContent : %s', index, JSON.stringify(jsonContent));

                var htmlText = generateHtmlTeamDetail(jsonContent);

                cards.push(
                    new builder.HeroCard(session)
                    .title("Team : " + jsonContent["TeamName"])
                    .subtitle("Manager : " + jsonContent["TeamManager"])
                    .text(htmlText)
                )
            });
            msg.attachments(cards);
            session.conversationData.stage = 1;
            session.send(msg);
            session.endDialog();
            session.endConversation();
        } else {
            next();
        }
    },
    function(session, results, next) {
        console.log('enter teaminfo water fall func#4');
        console.log('session.conversationData.stage : %d', session.conversationData.stage);
        if (session.conversationData.stage == 4) {
            console.log('wrong team info provided = %d', session.conversationData.retry);
            if (session.conversationData.retry < 3) {
                session.conversationData.retry++;
                session.conversationData.stage = 2;
                console.log('let user try again');
                //session.send("Sorry, the team info you input is either inexistent or too general, please try again.");   
                session.send("I don't know this team, can you try again with some other key words？");
            } else {
                console.log('3 trials failure, quit');
                session.send("Sorry, I cannot find the team you want to contact, please verify your input is correct");
                session.endConversation("Quit Conversation");
            }
        }
    }
]);

bot.dialog('postback', [
        function(session, args, next) {
            var postinfo = session.message.text;
            var arr = postinfo.split('.');
            console.log(arr);
            if (arr[0] == 'postback') {
                var ser_type = arr[1];
                var key = arr[2];
                var sub_key = arr[3];
                if (ser_type == 'teaminfo') {
                    var title = session.conversationData.teamInfoData[key]['TeamName'] + ' ' + sub_key;
                    var entries = session.conversationData.teamInfoData[key][sub_key].split(';');
                    var result = '';
                    entries.forEach(element => {
                        result += `${element} <br>`;
                    });
                    var msg = new builder.Message(session);
                    msg.attachments([
                        new builder.ThumbnailCard(session)
                        .title(title)
                        .text(result)
                    ]);
                }
            }
            session.send(msg);
            session.endDialog();
        }
    ])
    .triggerAction({
        matches: /^postback/i,
        onSelectAction: (session, args, next) => {
            // Add the help dialog to the dialog stack 
            // (override the default behavior of replacing the stack)
            session.beginDialog(args.action, args);
        }
    });

    bot.dialog('newfeature', [
        function(session) {
            console.log('enter new feature');
            session.send("More features are coming!!!<br /> Please contact jinjiez@microsoft.com if you have any ideas or suggestions.").endDialog();
            session.endConversation();
        }
    ]);

/*
function fake_query(ser_type, ser_key) {
    console.log("enter fake query func#");
    if (ser_type == "teaminfo") {
        if (ser_key == "iis") {
            var fs = require("fs");
            var contents = fs.readFileSync("fake-result.json");
            var jsonContent = JSON.parse(contents);
            return jsonContent;
        }
        return null;
    }
    return null;
}
*/


/*
var addUtterance = (config, jsonUtterance) => {
    try {
        // Add an utterance
        rq({
            uri: config.uri,
            method: 'POST',
            headers: {
                'Ocp-Apim-Subscription-Key': config.LUIS_authoringKey
            },
            json: true,
            body: jsonUtterance
        }).then(function(repos) {
            console.log("Add utterance returned");
            console.log(repos);
        });

        console.log("Add utterance done");

    } catch (err) {
        console.log(`Error adding utterance:  ${err.message} `);
        //throw err;
    }
}
*/

/*
bot.dialog('test', [
        function(session) {

            console.log("testing");

            rq('https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/2c8a694d-30b9-4203-a0f2-5d03815ccea4?subscription-key=c82d1549c7f145ceb7d832468201575d&verbose=true&timezoneOffset=0&q=/123')
                .then(function(repos) {
                    console.log(repos);
                })
                .catch(function(err) {
                    // Crawling failed...
                });

            luis_query(session, 'tech', 'call cannot be routed to call queue');
            luis_query(session, 'tech', '12345678');
        }
    ])
    .triggerAction({
        matches: /^test$/i,
        onSelectAction: (session, args, next) => {
            // Add the test dialog to the dialog stack 
            // (override the default behavior of replacing the stack)
            session.beginDialog(args.action, args);
        }
    });
*/

/*
function luis_query(session, ser_type, expr) {
    console.log("enter LUIS query func#");
    var uri = luisEndpoint + encodeURIComponent(expr);
    console.log("uri : %s", uri);
    var options = {
        uri: uri,
        headers: {
            'User-Agent': 'Request-Promise'
        },
        simple: false,
        json: true // Automatically parses the JSON string in the response
    };
    rq(options)
        .then(function(repos) {
            console.log(repos);
            var record = repos['topScoringIntent'];
            console.log('User has %d repos', Object.keys(record).length);
            Object.keys(record).forEach(function(key) {
                console.log('Key : ' + key + ', Value : ' + record[key]);
            });
            session.conversationData.stage++;
            session.conversationData.intent = record['intent'];
            session.beginDialog(ser_type, {
                intent: record['intent']
            });
        })
        .catch(function(err) {
            // Crawling failed...
        });
}
*/

/*
bot.dialog('techinfo', [
    function(session, args, next) {
        console.log('enter tech water fall func#1');
        console.log('session.conversationData.stage : %d', session.conversationData.stage);
        //session.send(args);
        if (session.conversationData.stage == 1) {
            builder.Prompts.text(session, "Please describe the issue you are facing.");
            session.conversationData.stage++;
        } else if (session.conversationData.stage >= 2 && session.conversationData.stage <= 5) {
            next();
        } else {
            console.log("sth wrong !  \n session.conversationData.stage = %d ", session.conversationData.stage);
        }
    },

    function(session, results, next) {
        console.log("session.message.text : %s", session.message.text);

        console.log('enter tech water fall func#2');
        console.log("results.response : %s ; session.message.text : %s", results.response, session.message.text);

        if (session.conversationData.stage == 2) {
            session.conversationData.techDescription = session.message.text;
            luis_query(session, "techinfo", session.message.text);
            session.endDialog();
        } else {
            next();
        }
    },

    function(session, results, next) {
        console.log("session.message.text : %s", session.message.text);

        console.log('enter tech water fall func#3');
        console.log("results.response : %s ; session.message.text : %s", results.response, session.message.text);

        console.log("stage %d", session.conversationData.stage);
        console.log("session.conversationData.intent %s", session.conversationData.intent);
        if (session.conversationData.stage == 3) {
            if (session.conversationData.intent == 'None') {
                session.send("Sorry, we don't understand the description. Please try describing it with another method");
                session.conversationData.stage = 2;
            } else {
                //data_query(session, 'techinfo', 'skype'); //for test as REST API not ready
                console.log("dataquery");
                data_query(session, "techinfo", session.conversationData.intent);
                session.conversationData.stage++;
                session.endDialog();
            }
        } else {
            next();
        }
    },

    function(session, results, next) {
        console.log("session.message.text : %s", session.message.text);

        console.log('enter tech water fall func#4');
        console.log("results.response : %s ; session.message.text : %s", results.response, session.message.text);
        console.log("stage %d", session.conversationData.stage);
        if (session.conversationData.stage == 4) {
            var records = session.conversationData.techref;

            var msg = new builder.Message(session);
            var cards = new Array();
            console.log(records);
            //let suggestionTxt="";

            records.forEach(function(record, index) {
                //suggestionTxt += (" <br> <a href=\"" + record['Suggestion'] +"\">" + record['Suggestion'] +"</a>") ;
                var imagePath;
                if (record['Techsuggestion'].startsWith('onenote')) {
                    imagePath = oneNoteImageUrl;
                } else {
                    imagePath = wikiImageUrl;
                }

                console.log(imagePath);

                cards.push(
                    new builder.ThumbnailCard(session)
                    .title(`Suggestion# ${index+1}`)
                    .text(`<a href="${record['Techsuggestion']}">${record['Techsuggestion']}</a>`)
                    .images([
                        builder.CardImage.create(session, imagePath)
                    ])
                );

                //suggestionTxt+=`<a href="${record['Suggestion']}">Suggestion# ${index+1}</a> <br>`;
            });

            var record0 = records[0]; //get record0 to fetch support team info and PG team info as all the records contain the same info for this
            var result = `<h2> Support Team </h2> ${record0['TeamName']} (Search in <b>TeamInfo</b> for more details) <br> <h2> PG Contact </h2> ${record0['PGTeam']} `;

            cards.push(
                new builder.ThumbnailCard(session)
                .title('You may also contact...')
                .text(result)
            );

            msg.attachments(cards);
            session.send(msg);

            //session.send(suggestionTxt);

            var title = "Is that what you need?";
            var msg = new builder.Message(session);
            msg.attachments([
                new builder.ThumbnailCard(session)
                .title(title)
                .buttons([
                    builder.CardAction.messageBack(session, '', 'Yes')
                    .text('Yes'),
                    builder.CardAction.messageBack(session, '', 'No')
                    .text('No'),
                ])
            ]);
            session.send(msg);
            session.conversationData.stage++;
        } else {
            next();
        }
    },

    function(session, results, next) {
        console.log("session.message.text : %s", session.message.text);
        console.log("results.response : %s", results.response);
        console.log('enter tech water fall func#5');
        console.log("results.response : %s ; session.message.text : %s", results.response, session.message.text);

        if (session.conversationData.stage == 5) {
            var res = session.message.text;
            if (res == 'Yes') {
                var jsonstr = constructJsonUtterance(session.conversationData.techDescription, session.conversationData.intent);
                console.log("jsonstr : %s", jsonstr);
                addUtterance(configAddUtterance, jsonstr);
                session.send("Thank you for your feedback! Have a nice day");
                session.endConversation("Quit Conversation");
            } else {
                console.log('Not a good answer, retry = %d', session.conversationData.retry);
                if (session.conversationData.retry < 3) {
                    session.conversationData.retry++;
                    session.conversationData.stage = 2;
                    session.send("Please try describing it in another way");
                } else {
                    session.send("Sorry, I am not able to provide you with the technical references you want. Please try later as I'm keeping improving myself. For now, please try to get the team contact for the consulting");
                    session.endConversation("Quit Conversation");
                }
            }
        }
    }
]);
*/