// Url to google spreadsheet containing saved information
//spreadsheetUrl = 'https://docs.google.com/spreadsheets/d/1xj-jryVPy9n1q0NPBCZxkW_yn1p2ZsmsRBjgMvc9wlw';  
//googleChatUrl = 'https://chat.google.com/';

// Helpful docs for SpreadsheetApp: https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet

const spr = SpreadsheetApp.openById('1xj-jryVPy9n1q0NPBCZxkW_yn1p2ZsmsRBjgMvc9wlw');
const threadSheet = spr.getSheetByName("Topics");
const remindSheet = spr.getSheetByName("Reminders");

/**
 * Responds to an ADDED_TO_SPACE event in Hangouts Chat.
 *
 * @param {Object} event the event object from Hangouts Chat
 */
function onAddToSpace(event) {
  let response = "";

  response = "Added to: " + event.space.displayName + ". I look forward to remembering things here.";

  if (event.message) {
    // Bot added through @mention.
    response = response + "\n" + COMMANDS.getHelp();
  }

  return { "text": response };
}

/**
 * Responds to a REMOVED_FROM_SPACE event in Hangouts Chat.
 *
 * @param {Object} event the event object from Hangouts Chat
 */
function onRemoveFromSpace(event) {
  console.info("Bot removed from ", event.space.name);
}

function onCardClick(event){
  Logger.log(JSON.stringify(event));

  if((typeof event.action !== 'undefined') && 
        (typeof event.action.actionMethodName !== 'undefined')){

    if(event.action.actionMethodName === "DISMISS_REMINDER" &&
          (typeof event.common !== 'undefined') && 
          (typeof event.common.parameters !== 'undefined') &&
          (typeof event.common.parameters.row !== 'undefined')){

      let row = event.common.parameters.row;
      remindSheet.deleteRow(row);

      return COMMANDS.getReminders(event);
    }

  }

  return { "text": "Unknown card action."};
}

/**
 * Responds to a MESSAGE event in Google Chat.
 *
 * @param {Object} event the event object from Google Chat
 */
 function onMessage(event) {
  Logger.log(JSON.stringify(event));
  let unknown = "Sorry, did not understand your request.\n\n" + COMMANDS.getHelp(event).text;

  if (event.message.slashCommand) {
    let func = SLASH_COMMANDS[event.message.slashCommand.commandId];

    if (func) {
      return COMMANDS[func](event);
    }
  } else {
    let textWords = event.message.argumentText.trim().split(/\s+/).slice(0);
    let command = textWords[0];
    let args = textWords.slice(1);

    let commandFunc = TEXT_COMMANDS[command];
    if (commandFunc) {
      return COMMANDS[commandFunc](event, command, args);
    }
  }

  return { "text": unknown };
}

/////////////////////////////////////////////////////////////////////
// Helper functions
/////////////////////////////////////////////////////////////////////
SLASH_COMMANDS = [ 
  'NOOP', 
  'listTopics', 
  'getHelp', 
  'slashAddToTopic', 
  'addReminder', 
  'getReminders'
];

TEXT_COMMANDS = {
  "set-topic": 'addToTopic',
  "clear-topics": 'clearTopic'
};

COMMANDS = {
  addReminder: (event) => {
    let topic = event.message.argumentText.trim();

    let who = []
    let sender = event.message.sender.displayName;
    if(event.message.annotations !== null || event.message.annotations.length > 0){
      let annos = event.message.annotations;
      for(let i=0; i<annos.length; i++){
        let a = annos[i];
        if(a.type == "USER_MENTION"){
          who.push(a.userMention.user.displayName);
        }
      }
    }

    // We probably have these cluttering up the beginning of the message, and need to trim them out.
    // NOTE: If a mention is NOT in the start of the message, don't include it in the list of reminder targets.
    whoNew = [];
    for(let i=0; i<who.length; i++){
      if(topic.startsWith(`@${who[i]}`)){
        let skipLen = 1+who[i].length;
        topic = topic.substring(skipLen).trim();
        whoNew.push(who[i]);
      }
    }
    who = whoNew;

    if(who.length < 1){
      who.push(sender);
    }

    let where = "direct message";
    let thread = "";
    if(event.message.space !== null){
      where = event.message.space.displayName;

      if (event.message.thread !== null){
        let spaceType = event.message.space.type.toLowerCase();
        let threadparts = event.message.thread.name.split('/');
        thread = `https://chat.google.com/${spaceType}/${threadparts[1]}/${threadparts[3]}`;
      }
    }

    if (topic === null || topic.length < 1 || where === null) {
      return {text: "Invalid /remind command. Usage: */remind [user-mention] <topic-words>*"};
    }

    let values = [];
    for(let i=0; i<who.length; i++){
      values.push([who[i], sender, topic, where, thread, HELPERS.secondsToDate(event.eventTime.seconds)]);
    }
    
    let lastRow = remindSheet.getLastRow();
    
    remindSheet.getRange(lastRow + 1, 1, values.length, values[0].length).setValues(values);
  
    return {text: "Reminder set!"};
  },

  getReminders: (event) => {
    let myName = event.message.sender.displayName;
    let actionResponse = "NEW_MESSAGE";
    if((typeof event.common !== 'undefined') && 
            (typeof event.common.parameters !== 'undefined') && 
            (typeof event.common.parameters.sender !== 'undefined') ){
      myName = event.common.parameters.sender;
      actionResponse = "UPDATE_MESSAGE";
    }

    let lastRow = remindSheet.getLastRow();
    if (lastRow > 1) {
      let values = remindSheet.getRange(2, 1, lastRow-1, 6).getValues();

      let by = {};
      for(let i=0; i<values.length; i++){
        let row = values[i];

        // only grab my own mentions.
        if(myName == row[0]){
          row.push(i+2);

          // gather them by sender
          if (!by[row[1]]) {
            by[row[1]] = [row];
          } else {
            by[row[1]].push(row);
          }
        }
      }

      let senders = Object.keys(by);
      if(senders.length > 0){
        let reminderCards = [];
        for(let i=0; i<senders.length; i++){
          let sender = senders[i];
          let rows = by[sender];

          for(let j=0; j<rows.length; j++){
            let row = rows[j];

            let cardButtons = [
              {
                textButton: {
                  text: "Dismiss",
                  onClick: {
                    action: {
                      actionMethodName: "DISMISS_REMINDER",
                      parameters: [
                        {
                          key: "row", 
                          value: `${row[row.length-1]}`
                        },
                        {
                          key: "sender",
                          value: sender
                        }
                      ]
                    }
                  }
                }
              }
            ];

            let card = {
              header: { title: row[3] },
              sections: [{
                widgets: [
                  {keyValue: {
                    content: row[2],
                    contentMultiline: true,
                    bottomLabel: `recorded at: ${new Date(row[5])}`
                  }},
                  {buttons: cardButtons}
              ]}]
            };

            if(row[4] !== ""){
              cardButtons.push({
                textButton: {
                  text: "Go to thread",
                  onClick: { openLink: { url: row[4] } }
                }
              });
            }

            reminderCards.push(card);
          }
        }

        let response = {
          actionResponse: actionResponse,
          cards: reminderCards
        };

        Logger.log(JSON.stringify(response));

        return response;
      }
    }
  
    return {text: "You have no reminders."};
  },

  listTopics: (event) => {
    let response = "Here are your saved topics:\n";
    let lastRow = threadSheet.getLastRow();
    if (lastRow < 2) {
      return "No topics saved.";
    }
  
    let values = threadSheet.getRange(2, 1, lastRow-1, 4).getValues();
    let bySubj = {}
    for(let i=0; i<values.length; i++) {
      let row=values[i];
      if (!bySubj[row[1]]) {
        bySubj[row[1]] = [row];
      } else {
        bySubj[row[1]].push(row);
      }
    }
  
    let topics = Object.keys(bySubj);
    for(let i=0; i<topics.length; i++) {
      let topic = topics[i];
      vs = bySubj[topic];
      response += `\n*${topic}*:`;
      for(let j=0; j<vs.length; j++){
        response += `\n-   ${vs[j][0]}\n      (on ${new Date(vs[j][2])} by ${vs[j][3]})`;
      }
    }
  
    return { text: response };
  },

  clearTopic: (event, command, args) => {
    let target = args.join(" ");
    if(target.length < 1){
      target = "ALL";
    }

    let lastRow = threadSheet.getLastRow();
    if (lastRow < 2) {
      return "No topics to clear.";
    }
  
    let topics = threadSheet.getRange(2, 1, lastRow-1, 2).getValues();
    let cleared = [];
    let toDelete = [];
    for (let i=0; i<topics.length; i++) {
      let entry = topics[i];
      if (target === "ALL" || target === entry[1]) {
        cleared.push(entry[0]);
        toDelete.push(i+2);
      }
    }

    toDelete.sort((a,b)=>a<b?1:-1);
    toDelete.forEach((row)=>{
      threadSheet.deleteRow(row);
    });

    return {text: "Cleared threads for topic '*" + target + "*':\n\n" + cleared.join("\n")};
  },

  slashAddToTopic: (event) => {
    let topic = event.message.argumentText.trim();
    let sender = event.message.sender.displayName;

    let url = null;
    if(event.message.thread !== null && event.message.space !== null){
      let spaceType = event.message.space.type.toLowerCase();
      let threadparts = event.message.thread.name.split('/');
      url = `https://chat.google.com/${spaceType}/${threadparts[1]}/${threadparts[3]}`;
    }

    if (topic === null || topic.length < 1 || url === null) {
      return {text: "Invalid /topic command. Usage: */topic <topic-words>*"};
    }

    let values = [[url, topic, HELPERS.secondsToDate(event.eventTime.seconds), sender]];
    let lastRow = threadSheet.getLastRow();
    threadSheet.getRange(lastRow + 1, 1, values.length, values[0].length).setValues(values);
  
    return {text: "Thread: " + url + " added to topic: *" + topic + "*"};
  },
  
  addToTopic: (event, command, args) => {
    let sender = event.message.sender.displayName;
    let topic = null;
    let url = null;
    if (args.length > 1) {
      url = args[0];
      topic = args.slice(1).join(" ");
    }

    if (topic === null || url === null) {
      return {text: "Invalid set-topic command. Usage: *set <thread-url> <topic-words>*"};
    }

    let values = [[url, topic, HELPERS.secondsToDate(event.eventTime.seconds), sender]];
    let lastRow = threadSheet.getLastRow();
    threadSheet.getRange(lastRow + 1, 1, values.length, values[0].length).setValues(values);
  
    return {text: "Thread: " + url + " added to topic: *" + topic + "*"};
  },
  
  getHelp: (event) => {
    return { text: "Available commands: \n"
      + "  */remind [<user-mention>] <message>* to set a reminder for you or someone else\n"
      + "  */reminders* to retrieve reminders stored for you (also clears your reminders)\n"
      + "  */topics* to retrieve the topic-to-thread-URLs mappings\n"
      + "  */topic <topic description>* to set the topic of the current thread\n"
      + "  */help* Print this help message.\n"
      + "\nNon-slash commands:\n"
      + "  *set-topic <URL> <topic-string>* to assign a topic to the given thread URL\n"
      + "  *clear-topics [<topic-string>]* Clear topics. If topic-string is given, only clear those\n"
    };
  }
};

HELPERS = {
  secondsToDate: (seconds) => {
    let u = new Date(seconds * 1000);
    return u.toISOString();
  }
};
