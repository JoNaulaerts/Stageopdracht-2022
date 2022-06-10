using System;
using System.IO;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using PostmarkDotNet;
using PostmarkDotNet.Model;

namespace WebJobsSDKSample
{
    
    public class Functions
    {
        public string connectionStringAzAD = "";

        // Start this function when the queue with the name "queue" recieves a message.
        public static async void ProcessQueueMessage([QueueTrigger("queue")] string message, ILogger logger)
        {
            logger.LogInformation("I recieved this message: " + message);

            if (message.Substring(0, 6) == "webjob")
            {
                // Extract the authentication token.
                var token = message.Substring(7, message.Length - 7);
                logger.LogInformation("token: " + token);
                string _token = token;

                // Authenticate the webjob to make requests to the Graph API.
                var graphClient = new GraphServiceClient(
                    new DelegateAuthenticationProvider(
                        requestMessage =>
                        {
                            // Append the access token to the request.
                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", _token);

                            // Get event times in the current time zone.
                            requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

                            return Task.CompletedTask;
                        }));

                // Get the total amount of users.
                var amountOfUsers = await graphClient.Users
                    .Request()
                    .GetAsync();

                // Get the amount of guest users.
                var amountOfGuests = await graphClient.Users
                    .Request()
                    .Filter("userType eq 'guest'")
                    .GetAsync();

                // Get the total amount of groups.
                var amountOfgroups = await graphClient.Groups
                    .Request()
                    .GetAsync();

                logger.LogInformation("Amount of users: " + amountOfUsers.Count);
                logger.LogInformation("Amount of guest users: " + amountOfGuests.Count);
                logger.LogInformation("Amount of groupq: " + amountOfgroups.Count);

                // Send an email asynchronously:
                var email = new PostmarkMessage()
                {
                    To = "jo.naulaerts@ventigrate.dev",
                    From = "jo.naulaerts@ventigrate.dev",
                    TrackOpens = true,
                    Subject = "Report from Azure",
                    TextBody = "Hello dear Postmark user.",
                    HtmlBody = "<h2>Report from Azure:</h2><p>The total amount of <strong>users</strong> is: <strong> " + amountOfUsers.Count + "</strong></p>" +
                    "<p>The amount of <strong>guest users </strong>is: <strong>" + amountOfGuests.Count + "</strong></p>" +
                    "<p>The total amount of <strong>groups </strong>is: <strong>" + amountOfgroups.Count + "</strong></p>",
                    MessageStream = "broadcast",
                    Tag = "New Year's Email Campaign",
                 };

                var client = new PostmarkClient("73a7839b-b42f-4abd-b15a-9b5815635713");
                var sendResult = await client.SendMessageAsync(email);

                if (sendResult.Status == PostmarkStatus.Success) { logger.LogInformation("Email successfully sent."); }
                else { logger.LogInformation("Something went wrong during the sending of this email."); }
            }
        }
    }
}
