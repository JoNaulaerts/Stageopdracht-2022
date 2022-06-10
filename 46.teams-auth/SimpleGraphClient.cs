using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.Bot.Schema;
using Microsoft.Graph;
using Microsoft.Graph.Core;
using TeamsAuth.Models;
using System.Web;

namespace Microsoft.BotBuilderSamples
{
    // This class is a wrapper for the Microsoft Graph API
    // See: https://developer.microsoft.com/en-us/graph
    public class SimpleGraphClient
    {
        private readonly string _token;

        public SimpleGraphClient(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                throw new ArgumentNullException(nameof(token));
            }

            _token = token;
        }

        // Get information about the user.
        public async Task<User> GetMeAsync()
        {
            var graphClient = GetAuthenticatedClient();
            var me = await graphClient.Me.Request().GetAsync();
            return me;
        }

        // Get an Authenticated Microsoft Graph client using the token issued to the user.
        private GraphServiceClient GetAuthenticatedClient()
        {
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
            return graphClient;
        }

        public async Task<GraphServiceClient> PostNewPasswordAsync(ManageUserClass userInput)
        {
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

            var currentPassword = userInput.currentPassword;

            var newPassword = userInput.newPassword;

            await graphClient.Me
                .ChangePassword(currentPassword, newPassword)
                .Request()
                .PostAsync();
            return graphClient;
        }

        public async Task<string> ResetPasswordAsync(User userInput)
        {
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

            //var password = Membership.GeneratePassword(Int32, Int32);
            string randomPassword = Guid.NewGuid().ToString("d").Substring(1, 10);

            var user = new User
            {
                PasswordProfile = new PasswordProfile
                {
                    ForceChangePasswordNextSignIn = true,
                    Password = randomPassword
                }
            };

            await graphClient.Users[userInput.Id]
                .Request()
                .UpdateAsync(user);
            return randomPassword;
        }

        public async Task<string> PostCreateUserAsync(ManageUserClass userInput)
        {
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

            string randomPassword = Guid.NewGuid().ToString("d").Substring(1, 10);

            var user = new User
            {
                AccountEnabled = userInput.accountEnabled,
                DisplayName = userInput.displayName,
                MailNickname = userInput.mailNickname,
                UserPrincipalName = userInput.userPrincipalName,
                PasswordProfile = new PasswordProfile
                {
                    ForceChangePasswordNextSignIn = true,
                    //Password = "changeMe!"
                    Password = randomPassword
                }
            };

            var response = await graphClient.Users
                           .Request()
                           .AddAsync(user);

            return user.PasswordProfile.Password;
        }

        public async Task<Group> GetGroupByDisplayNameAsync(string displayNameInput)
        {
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

            var response = await graphClient.Groups
                           .Request()
                           .Filter("displayName eq '" + displayNameInput + "'")
                           .GetAsync();

            return response[0];
        }

        public async Task<Group> PostUpdateCreateGroupAsync(ManageGroupClass groupInput, string manageModeInput)
        {
            GraphServiceClient graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    requestMessage =>
                    {
                        // Append the access token to the request.
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", _token);

                        // Get event times in the current time zone.
                        requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

                        return Task.CompletedTask;
                    }));
            var group = new Group
            {
                Description = groupInput.groupDescription,
                DisplayName = groupInput.groupDisplayName,
                GroupTypes = new List<String>() { "Unified" },
                MailEnabled = true,
                MailNickname = groupInput.groupMailNickname,
                SecurityEnabled = false,
            };

            if (manageModeInput == "create")
            {
                var returnGroup = await graphClient.Groups
                                .Request()
                                .AddAsync(group);
                Console.Write(returnGroup);

                group = returnGroup;

                // wait for 10 seconds to make sure the group exists before you add a team
                System.Threading.Thread.Sleep(10000);

                if ((bool)groupInput.addTeam)
                {                
                    var team = new Team
                    {
                        MemberSettings = new TeamMemberSettings
                        {
                            AllowCreatePrivateChannels = true,
                            AllowCreateUpdateChannels = true
                        },
                        MessagingSettings = new TeamMessagingSettings
                        {
                            AllowUserEditMessages = true,
                            AllowUserDeleteMessages = true
                        },
                        FunSettings = new TeamFunSettings
                        {
                            AllowGiphy = true,
                            GiphyContentRating = GiphyRatingType.Strict
                        }               
                    };
 
                await graphClient.Groups[returnGroup.Id].Team
                    .Request()
                    .PutAsync(team);
                }

        } else if (manageModeInput == "update")
            {
                await graphClient.Groups[groupInput.groupId]
                    .Request()
                    .UpdateAsync(group);

                if ((bool)groupInput.addTeam)
                {
                    var team = new Team
                    {
                        MemberSettings = new TeamMemberSettings
                        {
                            AllowCreatePrivateChannels = true,
                            AllowCreateUpdateChannels = true
                        },
                        MessagingSettings = new TeamMessagingSettings
                        {
                            AllowUserEditMessages = true,
                            AllowUserDeleteMessages = true
                        },
                        FunSettings = new TeamFunSettings
                        {
                            AllowGiphy = true,
                            GiphyContentRating = GiphyRatingType.Strict
                        }
                    };
                    await graphClient.Groups[groupInput.groupId].Team
                    .Request()
                    .PutAsync(team);
                }
            }
            return group;
        }

        public async Task<Group> GetExistingGroupAsync(string groupId)
        {
            GraphServiceClient graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    requestMessage =>
                    {
                        // Append the access token to the request.
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", _token);

                        // Get event times in the current time zone.
                        requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

                        return Task.CompletedTask;
                    }));

            var group = await graphClient.Groups[groupId]
                .Request()
                //.Select("displayName,description,mailNickname")
                .GetAsync();
            return group;
        }

        public async Task<GraphServiceClient> DeleteGroupAsync(string groupId)
        {
            GraphServiceClient graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    requestMessage =>
                    {
                        // Append the access token to the request.
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", _token);

                        // Get event times in the current time zone.
                        requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

                        return Task.CompletedTask;
                    }));

            await graphClient.Groups[groupId]
                .Request()
                .DeleteAsync();
            return graphClient;
        }

        public async Task<bool> GetTeamVerifyAsync(string groupId)
        {
            GraphServiceClient graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    requestMessage =>
                    {
                        // Append the access token to the request.
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", _token);

                        // Get event times in the current time zone.
                        requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

                        return Task.CompletedTask;
                    }));

            var team = await graphClient.Groups[groupId].Team
                .Request()
                .GetAsync();

            var teamExists = true;

            if (team != null)
            {
                teamExists = true;
            }
            else
            {
                teamExists = false;
            }

            return teamExists;
        }

        public async Task<GraphServiceClient> AddGroupOwnerAsync(string groupId, string ownerPrincipalName)
        {
            GraphServiceClient graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    requestMessage =>
                    {
                        // Append the access token to the request.
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", _token);

                        // Get event times in the current time zone.
                        requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

                        return Task.CompletedTask;
                    }));

            var ownerId = await graphClient.Users[ownerPrincipalName]
                    .Request()
                    .GetAsync();

            var directoryObject = new DirectoryObject
            {
                Id = ownerId.Id
            };

            await graphClient.Groups[groupId].Owners.References
                .Request()
                .AddAsync(directoryObject);


            return graphClient;
        }

        //public async Task<GraphServiceClient> SharePointFollowSiteAsync(string siteName, User me)
        //{
        //    GraphServiceClient graphClient = new GraphServiceClient(
        //        new DelegateAuthenticationProvider(
        //            requestMessage =>
        //            {
        //                // Append the access token to the request.
        //                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", _token);

        //                // Get event times in the current time zone.
        //                requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

        //                return Task.CompletedTask;
        //            }));
        //    //if (me.FollowedSites includes siteName)
        //    //{

        //    //}


        //    var queryOptions = new List<QueryOption>()
        //        {
        //            new QueryOption("search", siteName)
        //        };

        //    var sitesResult = await graphClient.Sites
        //        .Request(queryOptions)
        //        .GetAsync();


        //    var value = new List<Site>()
        //    {
        //        new Site
        //        {
        //            Id = sitesResult[0].Id
        //        },
        //    };

        //    await graphClient.Users[me.Id].FollowedSites
        //        .Add(value)
        //        .Request()
        //        .PostAsync();

        //    return graphClient;
        //}

        //public async Task<GraphServiceClient> SharePointUnfollowSiteAsync(string siteName, User me)
        //{
        //    GraphServiceClient graphClient = new GraphServiceClient(
        //        new DelegateAuthenticationProvider(
        //            requestMessage =>
        //            {
        //                // Append the access token to the request.
        //                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", _token);

        //                // Get event times in the current time zone.
        //                requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

        //                return Task.CompletedTask;
        //            }));
        //    //if (me.FollowedSites includes siteName)
        //    //{

        //    //}


        //    var queryOptions = new List<QueryOption>()
        //        {
        //            new QueryOption("search", siteName)
        //        };

        //    var sitesResult = await graphClient.Sites
        //        .Request(queryOptions)
        //        .GetAsync();


        //    //var value = new List<Site>()
        //    //{
        //    //    new Site
        //    //    {
        //    //        Id = sitesResult[0].Id
        //    //    },
        //    //};

        //    var value = new List<Site>()
        //        {
        //            new Site
        //            {
        //                Id = sitesResult[0].Id
        //            },
        //        };


        //    await graphClient.Users[me.Id].FollowedSites
        //        .Remove(value)
        //        .Request()
        //        .PostAsync();

        //    return graphClient;
        //}
    }
}
