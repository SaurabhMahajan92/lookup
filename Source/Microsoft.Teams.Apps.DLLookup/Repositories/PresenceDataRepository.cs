// <copyright file="PresenceDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Repositories
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.DLLookup.Helpers;
    using Microsoft.Teams.Apps.DLLookup.Helpers.Extentions;
    using Microsoft.Teams.Apps.DLLookup.Models;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// This class helps fetching user presence and contact data.
    /// </summary>
    public class PresenceDataRepository : IPresenceDataRepository
    {
        private const string EndpointGuid = "4034fc8f-d79f-45fc-8133-a607a4acaab2";
        private const string AppAgentObject = "AppAgentObject";
        private const string UcwaAutoDiscoveryUri = "https://webdir.online.lync.com/autodiscover/autodiscoverservice.svc/root";
        private const string MSGraphBatchAPI = "https://graph.microsoft.com/beta/$batch";
        private readonly List<string> onlinePresenceOptions = new List<string> { "BUSY", "DONOTDISTURB", "ONLINE" };

        private readonly IMemoryCache memoryCache;
        private readonly IConfiguration configuration;
        private readonly TelemetryClient telemetryClient;
        private readonly IProtectedApiCallHelper protectedApiCallHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="PresenceDataRepository"/> class.
        /// </summary>
        /// <param name="memoryCache">Singleton memory cache object.</param>
        /// <param name="configuration">Singleton instance of application configuration.</param>
        /// <param name="telemetryClient">Singleton TelemetryClient instance used to send telemetry to Azure application insights.</param>
        /// <param name="protectedApiCallHelper">Scoped protectedApiCallHelper instance used to call Web APIs.</param>
        public PresenceDataRepository(IMemoryCache memoryCache, IConfiguration configuration, TelemetryClient telemetryClient, IProtectedApiCallHelper protectedApiCallHelper)
        {
            this.memoryCache = memoryCache;
            this.configuration = configuration;
            this.telemetryClient = telemetryClient;
            this.protectedApiCallHelper = protectedApiCallHelper;
        }

        /// <summary>
        /// Get User presence details in a batch.
        /// </summary>
        /// <param name="peoplePresenceDataArray">Array of People Presence Data object used to get presence information.</param>
        /// <returns>People Presence Data model data filled with presence information.</returns>
        public async Task<List<PeoplePresenceData>> GetBatchUserPresenceAsync(PeoplePresenceData[] peoplePresenceDataArray)
        {
            List<PeoplePresenceData> peoplePresenceDataList = new List<PeoplePresenceData>();
            List<PeoplePresenceData> peoplePresenceDataListReturnValue = new List<PeoplePresenceData>();

            foreach (PeoplePresenceData member in peoplePresenceDataArray)
            {
                string userPrincipalName = member.UserPrincipalName.ToLower();
                if (!this.memoryCache.TryGetValue(userPrincipalName, out PeoplePresenceData peoplePresenceDataObj))
                {
                    peoplePresenceDataObj = new PeoplePresenceData()
                    {
                        UserPrincipalName = member.UserPrincipalName,
                        Id = member.Id,
                    };
                    peoplePresenceDataList.Add(peoplePresenceDataObj);
                }
                else
                {
                    peoplePresenceDataListReturnValue.Add(peoplePresenceDataObj);
                }
            }

            if (peoplePresenceDataList.Count > 0)
            {
                this.telemetryClient.TrackEvent($"GetBatchUserPresenceAsync. Getting presence from MS Graph. User Count : {peoplePresenceDataList.Count()}, users :{JsonConvert.SerializeObject(peoplePresenceDataList)}");

                string[] scopes = { "Presence.Read.All" };

                MemoryCacheEntryOptions options = new MemoryCacheEntryOptions
                {
                    AbsoluteExpirationRelativeToNow = TimeSpan.FromSeconds(Convert.ToInt32(this.configuration["CacheInterval"])), // cache will expire in 60 seconds or 1 minutes
                };

                var presenceBatches = peoplePresenceDataList.SplitList(19); // MS Graph batch limit is 20

                foreach (var presenceBatch in presenceBatches)
                {
                    try
                    {
                        List<MSGraphBatchRequest> allRequests = MSGraphBatchRequestCreator.CreateBatchRequestPayloadForPresence("/users/{0}/presence", "GET", presenceBatch);
                        MSGraphBatchRequestPayload payload = new MSGraphBatchRequestPayload()
                        {
                            Requests = allRequests,
                        };

                        List<MSGraphBatchResponse<dynamic>> responses = await this.protectedApiCallHelper.CallGraphApiPostOnBehalfOfUser(scopes, MSGraphBatchAPI, JsonConvert.SerializeObject(payload));

                        if (responses != null)
                        {
                            foreach (var presenceInfo in responses)
                            {
                                try
                                {
                                    string presenceStatus = presenceInfo.Body.availability;
                                    string userAADId = presenceInfo.Body.id;
                                    PeoplePresenceData peoplePresence = new PeoplePresenceData()
                                    {
                                        Availability = presenceStatus.ToLower() == "available" ? "Online" : presenceStatus,
                                        UserPrincipalName = presenceInfo.Id.ToLower(),
                                        Id = userAADId,
                                    };

                                    this.memoryCache.Set(peoplePresence.UserPrincipalName, peoplePresence, options);

                                    peoplePresenceDataListReturnValue.Add(peoplePresence);
                                }
                                catch (Exception ex)
                                {
                                    if (presenceInfo != null)
                                    {
                                        this.telemetryClient.TrackTrace($"GetBatchUserPresenceAsync. An error occurred: {ex.Message}. PeoplePresenceDataJObject : {JsonConvert.SerializeObject(presenceInfo)}", SeverityLevel.Error);
                                        this.telemetryClient.TrackException(ex);
                                    }
                                    else
                                    {
                                        this.telemetryClient.TrackTrace($"GetBatchUserPresenceAsync. An error occurred: {ex.Message} ", SeverityLevel.Error);
                                        this.telemetryClient.TrackException(ex);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (presenceBatch != null)
                        {
                            this.telemetryClient.TrackTrace($"GetBatchUserPresenceAsync. An error occurred: {ex.Message}. PresenceBatch : {JsonConvert.SerializeObject(presenceBatch)}", SeverityLevel.Error);
                            this.telemetryClient.TrackException(ex);
                        }
                        else
                        {
                            this.telemetryClient.TrackTrace($"GetBatchUserPresenceAsync. An error occurred: {ex.Message} ", SeverityLevel.Error);
                            this.telemetryClient.TrackException(ex);
                        }
                    }
                }
            }
            else
            {
                this.telemetryClient.TrackEvent($"GetBatchUserPresenceAsync. Presence of all users found in memory.");
            }

            return peoplePresenceDataListReturnValue;
        }

        /// <summary>
        /// Gets online members count in a distribution list.
        /// </summary>
        /// <param name="groupId">Distribution list id.</param>
        /// <returns><see cref="Task{TResult}"/>Online members count in distribution list.</returns>
        public async Task<int> GetDistributionListMembersOnlineCountAsync(string groupId)
        {
            int onlineMembersCount = 0;

            try
            {
                // Get members
                List<DistributionListMember> members = await this.GetMembersList(groupId);

                List<PeoplePresenceData> peoplePresenceDataList = new List<PeoplePresenceData>();

                foreach (DistributionListMember member in members)
                {
                    string userPrincipalName = member.UserPrincipalName.ToLower();
                    if (!this.memoryCache.TryGetValue(userPrincipalName, out PeoplePresenceData peoplePresenceDataObj))
                    {
                        peoplePresenceDataObj = new PeoplePresenceData()
                        {
                            UserPrincipalName = member.UserPrincipalName,
                            Id = member.Id,
                        };
                        peoplePresenceDataList.Add(peoplePresenceDataObj);
                    }
                    else
                    {
                        if (this.onlinePresenceOptions.Contains(peoplePresenceDataObj.Availability.ToUpper()))
                        {
                            onlineMembersCount++;
                        }
                    }
                }

                if (peoplePresenceDataList.Count > 0)
                {
                    this.telemetryClient.TrackEvent($"GetDistributionListMembersOnlineCountAsync. Getting presence from MS Graph. User Count : {peoplePresenceDataList.Count()}, users :{JsonConvert.SerializeObject(peoplePresenceDataList)}");

                    string[] scopes = { "Presence.Read.All" };

                    MemoryCacheEntryOptions options = new MemoryCacheEntryOptions
                    {
                        AbsoluteExpirationRelativeToNow = TimeSpan.FromSeconds(Convert.ToInt32(this.configuration["CacheInterval"])), // cache will expire in 60 seconds or 1 minutes
                    };

                    var presenceBatches = peoplePresenceDataList.SplitList(19); // MS Graph batch limit is 20

                    foreach (var presenceBatch in presenceBatches)
                    {
                        try
                        {
                            List<MSGraphBatchRequest> allRequests = MSGraphBatchRequestCreator.CreateBatchRequestPayloadForPresence("/users/{0}/presence", "GET", presenceBatch);
                            MSGraphBatchRequestPayload payload = new MSGraphBatchRequestPayload()
                            {
                                Requests = allRequests,
                            };

                            List<MSGraphBatchResponse<dynamic>> responses = await this.protectedApiCallHelper.CallGraphApiPostOnBehalfOfUser(scopes, MSGraphBatchAPI, JsonConvert.SerializeObject(payload));

                            if (responses != null)
                            {
                                foreach (var presenceInfo in responses)
                                {
                                    try
                                    {
                                        string presenceStatus = presenceInfo.Body.availability;
                                        string userAADId = presenceInfo.Body.id;
                                        PeoplePresenceData peoplePresence = new PeoplePresenceData()
                                        {
                                            Availability = presenceStatus.ToLower() == "available" ? "Online" : presenceStatus,
                                            UserPrincipalName = presenceInfo.Id.ToLower(),
                                            Id = userAADId,
                                        };

                                        this.memoryCache.Set(peoplePresence.UserPrincipalName, peoplePresence, options);

                                        if (this.onlinePresenceOptions.Contains(peoplePresence.Availability.ToUpper()))
                                        {
                                            onlineMembersCount++;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        if (presenceInfo != null)
                                        {
                                            this.telemetryClient.TrackTrace($"GetDistributionListMembersOnlineCountAsync. An error occurred: {ex.Message}. PeoplePresenceDataJObject : {JsonConvert.SerializeObject(presenceInfo)}", SeverityLevel.Error);
                                            this.telemetryClient.TrackException(ex);
                                        }
                                        else
                                        {
                                            this.telemetryClient.TrackTrace($"GetDistributionListMembersOnlineCountAsync. An error occurred: {ex.Message} ", SeverityLevel.Error);
                                            this.telemetryClient.TrackException(ex);
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            if (presenceBatch != null)
                            {
                                this.telemetryClient.TrackTrace($"GetDistributionListMembersOnlineCountAsync. An error occurred: {ex.Message}. PresenceBatch : {JsonConvert.SerializeObject(presenceBatch)}", SeverityLevel.Error);
                                this.telemetryClient.TrackException(ex);
                            }
                            else
                            {
                                this.telemetryClient.TrackTrace($"GetDistributionListMembersOnlineCountAsync. An error occurred: {ex.Message} ", SeverityLevel.Error);
                                this.telemetryClient.TrackException(ex);
                            }
                        }
                    }
                }
                else
                {
                    this.telemetryClient.TrackEvent($"Presence of all users in group found in memory. Group id : {groupId}");
                }
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace(ex.Message + "--------------" + ex.StackTrace);
            }

            return onlineMembersCount;
        }

        /// <summary>
        /// Gets distribution list members using group API.
        /// </summary>
        /// <param name="groupID">Distribution list id to get members list.</param>
        /// <returns>DistributionListMember data model.</returns>
        private async Task<List<DistributionListMember>> GetMembersList(string groupID)
        {
            string[] scopes = { "Group.Read.All" };
            string graphQuery = string.Format("https://graph.microsoft.com/v1.0/groups/{0}/members?$top=100&$select=id,displayName,jobTitle,mail,userPrincipalName,userType", groupID);
            List<DistributionListMember> distributionListMemberList = await this.protectedApiCallHelper.CallGraphApiOnBehalfOfUser<DistributionListMember>(scopes, graphQuery);

            return distributionListMemberList.Where(x => string.Equals(x.UserType, "member", StringComparison.OrdinalIgnoreCase) &&
                                                         string.Equals(x.Type, "#microsoft.graph.user", StringComparison.OrdinalIgnoreCase)).ToList();
        }

        /// <summary>
        /// Get group member details.
        /// </summary>
        /// <param name="groupMailAddress">Group mail address.</param>
        /// <returns>group member details.</returns>
        private async Task<List<DistributionList>> GetGroupInfo(string groupMailAddress)
        {
            string[] scopes = { "Group.Read.All" };
            string graphQuery = string.Format("https://graph.microsoft.com/v1.0/groups?$filter=mail eq '{0}'&$select=id,displayName,mail,mailNickname,mailEnabled", groupMailAddress);

            List<DistributionList> distributionList = await this.protectedApiCallHelper.CallGraphApiOnBehalfOfUser<DistributionList>(scopes, graphQuery);
            return distributionList;
        }
    }
}
