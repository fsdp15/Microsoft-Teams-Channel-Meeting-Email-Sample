using System;

namespace MeetingChannelBot.Models
{
    public class MeetingResponse
    {
        public System.Net.HttpStatusCode StatusCode { get; set; }
        public string MeetingLink { get; set; }
        public DateTimeOffset? StartDateTime { get; set; }
        public DateTimeOffset? EndDateTime { get; set; }
    }
}
