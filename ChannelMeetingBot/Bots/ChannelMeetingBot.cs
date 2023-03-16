using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Newtonsoft.Json.Linq;
using ChannelMeetingBot.Bots;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Text;
using MeetingChannelBot.Models;
using ChannelMeetingBot.Models;

namespace Microsoft.BotBuilderSamples
{
    public class ChannelMeetingBot : TeamsActivityHandler
    {
        public readonly IConfiguration _configuration;
        private readonly ProactiveAppIntallationHelper _helper = new ProactiveAppIntallationHelper();
        private readonly ConcurrentDictionary<string, ConversationReference> _conversationReferences;

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            // Instal the bot for each user
            // Schedule a meeting in a channel containing all members
            // Send the meeting in the channel
            // Send the meeting to each member individually
            // Send the meeting through email to each channel member

            turnContext.Activity.RemoveRecipientMention();
            var text = turnContext.Activity.Text.Trim().ToLower();

            if (text.Contains("scheduleChannel".ToLower()))
            {
                var result = await ScheduleMeetingInChannelScopeAsync(turnContext, cancellationToken);

                if (result.StatusCode == System.Net.HttpStatusCode.Created)
                {
                    string messagetoSend = $"I have successfuly created a meeting for all members of this channel!" +
                        $"\n\n Meeting Link: {result.MeetingLink}" +
                        $"\n\n Start Datetime: {result.StartDateTime}" +
                        $"\n\n End Datetime: {result.EndDateTime}";
                    var emailSent = await SendEmailToAllMembers(turnContext, messagetoSend, cancellationToken );

                    // Append email message
                    if (emailSent == true)
                    {
                        string emailMessagetoSend = messagetoSend + $"\n\n I have also sent an e-mail to each of you with the meeting details!";
                        string chatMessagetoSend = messagetoSend + $"\n\n I have also sent an e-mail to you with the meeting details!";
                        await turnContext.SendActivityAsync(MessageFactory.Text(emailMessagetoSend), cancellationToken);

                        await SendNotificationToAllUsersAsync(turnContext, chatMessagetoSend, cancellationToken );
                    }
                    else
                    {
                        await turnContext.SendActivityAsync(MessageFactory.Text($"Failure while setting up a meeting"), cancellationToken);
                    }                                                      
                }
                else
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text($"Failure while setting up a meeting"), cancellationToken);
                }
            }
            else if (text.Contains("install".ToLower()))
            {
                var result = await InstalledAppsinPersonalScopeAsync(turnContext, cancellationToken);
                await turnContext.SendActivityAsync(MessageFactory.Text($"Existing: {result.Existing} \n\n Newly Installed: {result.New}"), cancellationToken);
            }
            else
            {
                await CardActivityAsync(turnContext, false, cancellationToken);
            }
        }

        public ChannelMeetingBot(ConcurrentDictionary<string, ConversationReference> conversationReferences, IConfiguration configuration)
        {
            _conversationReferences = conversationReferences;
            _configuration = configuration;
        }

        private void AddConversationReference(Activity activity)
        {
            var conversationReference = activity.GetConversationReference();
            _conversationReferences.AddOrUpdate(conversationReference.User.AadObjectId, conversationReference, (key, newValue) => conversationReference);
        }

        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            foreach (var member in membersAdded)
            {
                if (member.Id != turnContext.Activity.Recipient.Id)
                {
                    // Add current user to conversation reference.
                    AddConversationReference(turnContext.Activity as Activity);
                }
            }
        }

        public async Task<MeetingResponse> ScheduleMeetingInChannelScopeAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var currentPage_Members = await TeamsInfo.GetPagedMembersAsync(turnContext, null, null, cancellationToken);

            if (BotIsInstalledForAllUsers(currentPage_Members.Members.Count) == false)
            {
                // send message and end this method
                await turnContext.SendActivityAsync(MessageFactory.Text($"The bot is not installed for all members of this channel! Please click on the Install the bot for all users" +
                    $" button first"), cancellationToken);
                return new MeetingResponse
                {
                    StatusCode = System.Net.HttpStatusCode.InternalServerError
                };
            }

            TeamDetails teamDetails = await TeamsInfo.GetTeamDetailsAsync(turnContext, turnContext.Activity.TeamsGetTeamInfo().Id, cancellationToken);
            if (teamDetails != null)
            {
                // Schedules a 30 min meeting for 30 min ahead
                string Access_Token = await _helper.GetToken(turnContext.Activity.Conversation.TenantId, _configuration["MicrosoftAppId"], _configuration["MicrosoftAppPassword"]);
                GraphServiceClient graphClient = _helper.GetAuthenticatedClient(Access_Token);

                var owners = await graphClient.Groups[teamDetails.AadGroupId].Owners.Request().GetAsync();

                if (owners.Count == 0)
                {
                    return new MeetingResponse
                    {
                        StatusCode = System.Net.HttpStatusCode.InternalServerError
                    };
                    // Team must have an owner
                }
                
                Graph.MeetingParticipantInfo organizer = new Graph.MeetingParticipantInfo
                {
                    Identity = new IdentitySet
                    {
                        User = new Microsoft.Graph.Identity
                        {
                            Id = owners.CurrentPage[0].Id
                        }
                    }
                };

                List<Graph.MeetingParticipantInfo> attendees = new List<Graph.MeetingParticipantInfo>();
                foreach (var teamMember in currentPage_Members.Members)
                {
                    attendees.Add(new Graph.MeetingParticipantInfo
                    {
                        Identity = new IdentitySet
                        {
                            User = new Microsoft.Graph.Identity
                            {
                                Id = teamMember.AadObjectId
                            }
                        }
                    });
                }

                MeetingParticipants participants = new MeetingParticipants
                {
                   Attendees = attendees,
                   Organizer = organizer
                };

                var requestBody = new OnlineMeeting
                {
                    StartDateTime = DateTime.Now.AddMinutes(30),
                    EndDateTime = DateTime.Now.AddMinutes(60),
                    Subject = "Graph API Channel Meeting",
                    Participants = participants
                };

                /* I couldn't figure out how to call the onlineMeetings API with application permissions
                    through Graph API C# SDK. The documentation only provides examples for delegated:
                    https://learn.microsoft.com/en-us/graph/api/application-post-onlinemeetings?view=graph-rest-1.0&tabs=http
                    Therefore, I resorted to using a default C# HTTP Call. The same happened with the e-mail API.
                */

                using (HttpClient c = new HttpClient())
                {
                    string url = "https://graph.microsoft.com/v1.0/users/" +
                        owners.CurrentPage[0].Id + "/onlineMeetings";

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, url);

                    HttpContent httpContent = new StringContent(System.Text.Json.JsonSerializer.Serialize(requestBody), Encoding.UTF8, "application/json");

                    request.Content = httpContent;
                    //Authentication token
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", Access_Token);


                    var response = await c.SendAsync(request, cancellationToken);
                    var responseString = await response.Content.ReadAsStringAsync(cancellationToken);

                    OnlineMeeting onlineMeeting = System.Text.Json.JsonSerializer.Deserialize<OnlineMeeting>(responseString);

                    if (response.StatusCode == System.Net.HttpStatusCode.Created)
                    {
                        return new MeetingResponse
                        {
                            StatusCode = System.Net.HttpStatusCode.Created,
                            MeetingLink = onlineMeeting.AdditionalData["joinUrl"].ToString(),
                            StartDateTime = onlineMeeting.StartDateTime,
                            EndDateTime = onlineMeeting.EndDateTime
                        };
                    }
                    else
                    {
                        return new MeetingResponse
                        {
                            StatusCode = System.Net.HttpStatusCode.InternalServerError
                        };
                    }
                }        
            }
            else
            {
                return new MeetingResponse
                {
                    StatusCode = System.Net.HttpStatusCode.InternalServerError
                };
            }
        }

        public async Task<Boolean> SendEmailToAllMembers(ITurnContext<IMessageActivity> turnContext, String messagetoSend,  CancellationToken cancellationToken)
        {
            var currentPage_Members = await TeamsInfo.GetPagedMembersAsync(turnContext, null, null, cancellationToken);

            if (BotIsInstalledForAllUsers(currentPage_Members.Members.Count) == false)
            {
                // send message and end this method
                await turnContext.SendActivityAsync(MessageFactory.Text($"The bot is not installed for all members of this channel! Please click on the Install the bot for all users" +
                    $" button first"), cancellationToken);
                return false;
            }

            TeamDetails teamDetails = await TeamsInfo.GetTeamDetailsAsync(turnContext, turnContext.Activity.TeamsGetTeamInfo().Id, cancellationToken);

            if (teamDetails != null)
            {
                string Access_Token = await _helper.GetToken(turnContext.Activity.Conversation.TenantId, _configuration["MicrosoftAppId"], _configuration["MicrosoftAppPassword"]);
                GraphServiceClient graphClient = _helper.GetAuthenticatedClient(Access_Token);

                var owners = await graphClient.Groups[teamDetails.AadGroupId].Owners.Request().GetAsync();

                if (owners.Count == 0)
                {
                    return false;
                    // Team must have an owner
                }

                Microsoft.Graph.User owner = (Microsoft.Graph.User)(owners.CurrentPage[0]);
                string ownerEmail = owner.UserPrincipalName;

                List<Recipient> recipients = new List<Recipient>();
                foreach (var teamMember in currentPage_Members.Members)
                {
                    recipients.Add(new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = teamMember.Email,
                        },
                    });
                }

                var message = new Message
                    {
                        Subject = "You have been invited to a channel meeting!",
                        Body = new ItemBody
                        {
                            ContentType = BodyType.Text,
                            Content = messagetoSend,
                        },
                        Sender = new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = ownerEmail, // Sending email as the channel's owner
                            },
                        },
                        From = new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = ownerEmail,
                            },
                        },
                        ToRecipients = recipients
                };
                message.ODataType = null; // The API does not accept this parameter in the JSON body.

                var requestPayload = new MailContentModel
                {
                    message = message
                };

                using (HttpClient c = new HttpClient())
                {
                    string url = "https://graph.microsoft.com/v1.0/users/" +
                        owners.CurrentPage[0].Id + "/sendMail";

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, url);

                    HttpContent httpContent = new StringContent(System.Text.Json.JsonSerializer.Serialize(requestPayload), 
                        Encoding.UTF8, "application/json");

                    request.Content = httpContent;
                    //Authentication token
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", Access_Token);

                    var response = await c.SendAsync(request, cancellationToken);
                    var responseString = await response.Content.ReadAsStringAsync(cancellationToken);

                    if (response.StatusCode == System.Net.HttpStatusCode.Accepted)
                    {
                        return true;
                    }
                    else
                    {

                        return false;
                    }
                }
            }
                return false;
        }

        public async Task SendNotificationToAllUsersAsync(ITurnContext<IMessageActivity> turnContext, string message, CancellationToken cancellationToken)
        {
            var currentPage_Members = await TeamsInfo.GetPagedMembersAsync(turnContext, null, null, cancellationToken);

            if (BotIsInstalledForAllUsers(currentPage_Members.Members.Count) == false)
            {
                // send message and end this method
                await turnContext.SendActivityAsync(MessageFactory.Text($"The bot is not installed for all members of this channel! Please click on the Install the bot for all users" +
                    $" button first"), cancellationToken);
            }

            // Send notification to all the members
            foreach (var conversationReference in _conversationReferences.Values)
            {
                await turnContext.Adapter.ContinueConversationAsync(_configuration["MicrosoftAppId"], conversationReference,
                                                       async (t2, c2) =>
                                                       {
                                                           await t2.SendActivityAsync(MessageFactory.Text(message), c2).ConfigureAwait(false);
                                                       }, cancellationToken);
            }
        }

        public bool BotIsInstalledForAllUsers(int membersCount)
        {
            if (_conversationReferences.Count < membersCount)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public async Task<InstallationCounts> InstalledAppsinPersonalScopeAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var currentPage_Members = await TeamsInfo.GetPagedMembersAsync(turnContext, null, null, cancellationToken);
            int existingAppInstallCount = _conversationReferences.Count;
            int newInstallationCount = 0;

            foreach (var teamMember in currentPage_Members.Members)
            {
                // Check if present in App Conversation reference
                if (!_conversationReferences.ContainsKey(teamMember.AadObjectId))
                {
                    // Perform installation for all the member whose conversation reference is not available.
                    await _helper.AppinstallationforPersonal(teamMember.AadObjectId, turnContext.Activity.Conversation.TenantId, _configuration["MicrosoftAppId"], _configuration["MicrosoftAppPassword"], _configuration["AppCatalogTeamAppId"]);
                    newInstallationCount++;
                }
            }

            return new InstallationCounts
            {
                
            Existing = existingAppInstallCount,
                New = newInstallationCount
            };
        }

        private static async Task CardActivityAsync(ITurnContext<IMessageActivity> turnContext, bool update, CancellationToken cancellationToken)
        {

            var card = new HeroCard
            {
                Buttons = new List<CardAction>
                        {
                            new CardAction
                            {
                                Type = ActionTypes.MessageBack,
                                Title = "Schedule Channel Meeting",
                                Text = "scheduleChannel"
                            },
                            new CardAction
                            {
                                Type = ActionTypes.MessageBack,
                                Title = "Install the bot for all users",
                                Text = "install"
                            }
                        }
            };

            await SendWelcomeCard(turnContext, card, cancellationToken);

        }

        private static async Task SendWelcomeCard(ITurnContext<IMessageActivity> turnContext, HeroCard card, CancellationToken cancellationToken)
        {
            _ = new JObject { { "count", 0 } };
            card.Title = "Welcome!";

            var activity = MessageFactory.Attachment(card.ToAttachment());

            await turnContext.SendActivityAsync(activity, cancellationToken);
        }
    }
}