# Microsoft-Teams-Channel-Meeting-Email-Sample
A Microsoft Teams Bot sample that can schedule meetings for a team and also send emails to its users.

Hi!

This app is a bot for Microsoft Teams, made for the MS Graph Hack Together Hackathon: https://github.com/microsoft/hack-together.

It is a proof of concept showing how to schedule a meeting for a team in Microsoft Teams automatically by a bot using
MS Graph API, and send an email to the users with the invite link. I chose this idea because I haven't found any MS Teams
Bot Sample similar to this in https://learn.microsoft.com/en-us/samples/browse/?products=office-teams.

The bot first install itself for all the users of the team, and for this, I used the template Graph Proactive Installation with some modifications: https://github.com/OfficeDev/Microsoft-Teams-Samples/tree/main/samples/graph-proactive-installation/csharp

Later on, adding to the code above, the bot presents the option for the user to schedule the meeting. A meeting link is then generated, with all of the Team members invited to the meeting. The bot posts the meeting link into the channel he is installed.
After that, the bot also sends an e-mail to each user in the channel containing the meeting link. Finally, the bot
also proactively sends a 1:1 message to each user in the channel containing the meeting link.

There is a README.docx file in the repository showcasing how to install the bot.

Since the installation can be quite long/difficult and prone to errors, I created a video to showcase the bot in case
you do not manage to get it working: https://www.youtube.com/watch?v=jLOSrWPgOpU. Mirror: https://1drv.ms/v/s!AhbdjVk58i1ahdl3X8v1SDxEsRitBw?e=ljT46t.
