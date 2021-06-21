// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, MessageFactory } = require('botbuilder');

class EchoBot extends ActivityHandler {
  constructor() {
    super();
    // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
    this.onMessage(async (context, next) => {
      switch (context.activity.text) {
        case 'invoke':
          await context.sendActivity(
            MessageFactory.attachment(
              {
                content: {
                  type: 'AdaptiveCard',
                  body: [
                    {
                      type: 'TextBlock',
                      text: 'Select an `Action.Execute` action:'
                    },
                    {
                      type: 'ActionSet',
                      actions: [
                        {
                          type: 'Action.Execute',
                          verb: 'dump',
                          title: 'Dump activity',
                          data: {
                            hello: 'World!'
                          }
                        }
                      ]
                    }
                  ],
                  $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
                  // version: '2.3'
                  version: '1.4'
                },
                contentType: 'application/vnd.microsoft.card.adaptive'
              },
              'Showing a card'
            )
          );

          break;

        default:
          const replyText = `Echo: ${context.activity.text}`;

          await context.sendActivity(MessageFactory.text(replyText, replyText));
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      const welcomeText = 'Hello and welcome!';
      for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
        if (membersAdded[cnt].id !== context.activity.recipient.id) {
          await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
        }
      }
      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onInvokeActivity = async context => {
      console.log({
        activity: context.activity
      });

      const { activity } = context;

      console.log(activity.value.verb);

      switch (activity.value.verb) {
        case 'dump':
          return ActivityHandler.createInvokeResponse(
            adaptiveCardResponse({
              type: 'AdaptiveCard',
              body: [
                {
                  text: 'Dump activity',
                  type: 'TextBlock'
                },
                {
                  text: '```\n' + JSON.stringify(activity, null, 2) + '\n```',
                  type: 'TextBlock'
                }
              ],
              $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
              version: '1.4'
            })
          );

        default:
          return ActivityHandler.createInvokeResponse({
            type: 'application/vnd.microsoft.activity.message',
            value: 'Done'
          });
      }
    };

    // this.onAdaptiveCardInvoke = async (context, invokeValue) => {
    //   console.log(context);
    //   console.log(invokeValue);

    //   return {};
    // };
  }
}

function adaptiveCardResponse(cardContent) {
  return {
    // statusCode: 200,
    type: 'application/vnd.microsoft.card.adaptive',
    value: cardContent
  };
}

module.exports.EchoBot = EchoBot;
