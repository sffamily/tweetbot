const notificationTemplate = require("./adaptiveCards/notification-default.json");
const { AdaptiveCards } = require("@microsoft/adaptivecards-tools");
const { bot } = require("./internal/initialize");
const needle = require("needle");
const teamsfxSdk = require("@microsoft/teamsfx");

// Load application configuration
const teamsFx = new teamsfxSdk.TeamsFx();

const twtapi_endpoint = teamsFx.getConfig("TWTAPI_ENDPOINT");
const twtapi_bearer_token = teamsFx.getConfig("TWTAPI_BEARER_TOKEN");

//let last_id = BigInt('1526537624019881984');
let dt = new Date("2022-05-18T00:00:00.000Z");
// Time trigger to send notification. You can change the schedule in ../timerNotifyTrigger/function.json
module.exports = async function (context, myTimer) {
  dt.setHours(dt.getHours() - 1);
  const timeStamp = dt.toISOString();
  dt = new Date();
  for (const target of await bot.notification.installations()) {
    const params = {
      'query': '"teams toolkit" -is:retweet',
      'tweet.fields': 'id,created_at,text,author_id',
      //'since_id': last_id,
      'start_time': timeStamp,
    }
    const response = await needle('get', twtapi_endpoint + '/2/tweets/search/recent', params, {
      headers: {
        'Authorization': 'Bearer ' + twtapi_bearer_token,
      }
    });

    if(response.body.data) {
      for (const tweet of response.body.data) {
        //if (BigInt(tweet.id) > last_id) {
        //  last_id = BigInt(tweet.id);
        //}
        const params_user = {
          'user.fields': 'name,username,profile_image_url',
        }
        const response_user = await needle('get', twtapi_endpoint + '/2/users/' + tweet.author_id, params_user, {
          headers: {
            'Authorization': 'Bearer ' + twtapi_bearer_token,
          }
        });

        await target.sendAdaptiveCard(
          AdaptiveCards.declare(notificationTemplate).render({
            tweetuser: response_user.body.data.name,
            tweetusername: response_user.body.data.name+" @"+response_user.body.data.username,
            tweetprofile: response_user.body.data.profile_image_url,
            tweettext: tweet.text,
            tweettime: tweet.created_at,
            tweeturl: "https://twitter.com/" + response_user.body.data.username + "/status/" + tweet.id,
          })
        );
      };
    }
    else {
      //await target.sendMessage("No new tweets since last check with tweet_id: " + last_id);
      await target.sendMessage("No new tweets since: " + timeStamp);
    }
  }

  /****** To distinguish different target types ******/
  /** "Channel" means this bot is installed to a Team (default to notify General channel)
  if (target.type === "Channel") {
    // Directly notify the Team (to the default General channel)
    await target.sendAdaptiveCard(...);
    // List all channels in the Team then notify each channel
    const channels = await target.channels();
    for (const channel of channels) {
      await channel.sendAdaptiveCard(...);
    }
    // List all members in the Team then notify each member
    const members = await target.members();
    for (const member of members) {
      await member.sendAdaptiveCard(...);
    }
  }
  **/

  /** "Group" means this bot is installed to a Group Chat
  if (target.type === "Group") {
    // Directly notify the Group Chat
    await target.sendAdaptiveCard(...);
    // List all members in the Group Chat then notify each member
    const members = await target.members();
    for (const member of members) {
      await member.sendAdaptiveCard(...);
    }
  }
  **/

  /**  "Person" means this bot is installed as a Personal app
  if (target.type === "Person") {
    // Directly notify the individual person
    await target.sendAdaptiveCard(...);
  }
  **/
};
