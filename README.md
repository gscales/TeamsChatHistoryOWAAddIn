# Microsoft Teams Chat History Outlook/OWA Addin #

Being somebody who is transitioning across from Skype for Business to Teams one of things I missed the most (and found the most frustrating) is the lack of the ability in Outlook and OWA to view the conversation history from Online meetings and private chats in Microsoft Teams. This is especially frustrating when you have an external meeting and your sent an IM that contains some vital information for what you need to do. This information is tracked in your mailbox for compliance reasons in the TeamsMessagesData  folder but this folder is hidden so it not accessible to the clients and must be extracted by other means eg. Most people seem to point to doing a compliance search if you need this data [https://docs.microsoft.com/en-us/microsoftteams/security-compliance-overview](https://docs.microsoft.com/en-us/microsoftteams/security-compliance-overview) .

Given that the information is in my mailbox and there shouldn't be any privacy concern arounds access EWS (Exchange Web Services) can be used to both find the TeamsMessagesData folder in the Non_IPM_Subtree of the Mailbox and then FindItems is used to return the Chat compliance messages. The properties used in the display are the BodyPreivew and WebReadLink which can be used to open the Message in OWA


One thing I did find after using this for a while is that it didn't work when I got a notification from teams like the following



Because the above search came from noreply@email.teams.microsoft.com it couldn't be used in the above query. Looking at the notification message unfortunately there wasn't any other properties that did contain the email address but the full displayname of the user was used in the email's displayName so as a quick workaround for these I made use of EWS's resolvenames operation to resolve the displayName to an email address and then I could use the Addin even on the notification messages to see the private chat message that was sent to me within OWA without needing to open the Teams app (which if you have multiple tennants can be a real pain). So this one turned into a real productiviy enhancer for me.

Want to give it a try yourself ?

I've hosted the files on my GitHub pages so its easy to test (if you like it clone it and host it somewhere else). But all you need to do is add it as a custom addin (if you allowed to) using the 
URL-

  https://gscales.github.io/TeamsChatHistory/TeamsChatHistory.xml

![](https://1.bp.blogspot.com/-fV2Wxo0Cr7Q/XIHvqow9SVI/AAAAAAAACTM/Vo-Pc3Q74AgZ-_KNh_Nl9UhZmBbouJS1wCLcBGAs/s1600/addin3.JPG)



