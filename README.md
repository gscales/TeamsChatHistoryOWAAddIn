# Microsoft Teams Chat History Outlook/OWA Addin #

Being somebody who is transitioning across from Skype for Business to Teams one of things I missed the most (and found the most frustrating) is the lack of the ability in Outlook and OWA to view the conversation history from Online meetings and private chats in Microsoft Teams. This is especially frustrating when you have an external meeting and your sent an IM that contains some vital information for what you need to do. This information is tracked in your mailbox for compliance reasons in the Teams Chat folder but this folder is hidden so it not accessible to the clients and must be extracted by other means eg. Most people seem to point to doing a compliance search if you need this data [https://docs.microsoft.com/en-us/microsoftteams/security-compliance-overview](https://docs.microsoft.com/en-us/microsoftteams/security-compliance-overview) .

Given that the information is in my mailbox and there shouldn't be any privacy concern arounds access it I looked at few ways of getting access to these TeamsChat messages in OWA and Outlook the first way was using a SearchFolder. This did kind of work but because of a few quirks that the Hidden folder caused was only usable when using Outlook in online mode (which isn't very usable). The next thing I did was look a using an Addin which worked surprising well and was relatively easy to implement. Here is what it looks likes in action all you need to do is find a Message from the user you want to view the chat message from and then a query will be executed to find the Chat messages from that user using the Outlook REST endpoint eg

![](https://1.bp.blogspot.com/-j_fyvriDXUQ/XIHgQL1X3cI/AAAAAAAACS0/nJKkqlPsjdIJpNohzk6p9Mi_DmmZwU3LACLcBGAs/s1600/tcHist1.JPG)


That constructs a query that look like the following to Outlook REST endpoint

    https://outlook.office.com/api/v2.0/me/MailFolders/AllItems/messages?$Top=100&amp;$Select=ReceivedDateTime,bodyPreview,weblink&amp;$filter=SingleValueExtendedProperties/Any(ep: ep/PropertyId eq 'String 0x001a' and ep/Value eq 'IPM.SkypeTeams.Message') and from/emailAddress/address eq 'e5tmp5@domain.com' 

To break this down a bit first this gets the first 100 messages from the AllItems Search Folder

 https://outlook.office.com/api/v2.0/me/MailFolders/AllItems/messages?$Top=100
Next this selects the properties we are going to use the table to display, I used body preview because for IM's that generally don't have subjects so getting the body preview text is generally good enough to shown the whole message. But if the message is longer the link is provided which will open up in a new OWA windows using the weblink property which contains a full path to open the Item. One useful things about opening the message this way is you can then click replay and continue a message from IM in email with the body context from the IM (I know this will really erk some Teams people but i think it pretty cool and has proven useful for me).

    $Select=ReceivedDateTime,bodyPreview,weblink

Next this is the filter that is applied so it only returns the IM chat message (or those messages that have an ItemClass of IPM.SkypeTeams.Message and are from the sender assoicated with the Message you activate the Addin on.

 `$filter=SingleValueExtendedProperties/Any(ep: ep/PropertyId eq 'String 0x001a' and ep/Value eq 'IPM.SkypeTeams.Message') and from/emailAddress/address eq 'e5tmp5@datarumble.com'`

One thing I did find after using this for a while is that it didn't work when I got a notification from teams like the following



Because the above search came from noreply@email.teams.microsoft.com it couldn't be used in the above query. Looking at the notification message unfortunately there wasn't any other properties that did contain the email address but the full displayname of the user was used in the email's displayName so as a quick workaround for these I made use of EWS's resolvenames operation to resolve the displayName to an email address and then I could use the Addin even on the notification messages to see the private chat message that was sent to me within OWA without needing to open the Teams app (which if you have multiple tennants can be a real pain). So this one turned into a real productiviy enhancer for me.

Want to give it a try yourself ?

I've hosted the files on my GitHub pages so its easy to test (if you like it clone it and host it somewhere else). But all you need to do is add it as a custom addin (if you allowed to) using the 
URL-

  https://gscales.github.io/TeamsChatHistory/TeamsChatHistory.xml

![](https://1.bp.blogspot.com/-fV2Wxo0Cr7Q/XIHvqow9SVI/AAAAAAAACTM/Vo-Pc3Q74AgZ-_KNh_Nl9UhZmBbouJS1wCLcBGAs/s1600/addin3.JPG)



