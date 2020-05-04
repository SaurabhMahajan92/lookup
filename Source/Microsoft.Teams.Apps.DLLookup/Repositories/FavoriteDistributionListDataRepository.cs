// <copyright file="FavoriteDistributionListDataRepository.cs" company="Microsoft">
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
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.DLLookup.Helpers;
    using Microsoft.Teams.Apps.DLLookup.Helpers.Extentions;
    using Microsoft.Teams.Apps.DLLookup.Models;
    using Newtonsoft.Json;

    /// <summary>
    /// This class contains read/write operations for distribution list member data on AAD and table storage.
    /// </summary>
    public class FavoriteDistributionListDataRepository : FavoriteDistributionListTableRepository
    {
        private readonly TelemetryClient telemetryClient;
        private readonly IProtectedApiCallHelper protectedApiCallHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="FavoriteDistributionListDataRepository"/> class.
        /// </summary>
        /// <param name="configuration">Singleton instance of application configuration.</param>
        /// <param name="telemetryClient">Singleton TelemetryClient instance used to send telemetry to Azure application insights.</param>
        /// <param name="protectedApiCallHelper">Scoped protectedApiCallHelper instance used to call Web APIs.</param>
        public FavoriteDistributionListDataRepository(
            IConfiguration configuration, TelemetryClient telemetryClient, IProtectedApiCallHelper protectedApiCallHelper)
            : base(
                configuration)
        {
            this.telemetryClient = telemetryClient;
            this.protectedApiCallHelper = protectedApiCallHelper;
        }

        /// <summary>
        /// Creates/Updates favorite distribution list data in table storage.
        /// </summary>
        /// <param name="favoriteDistributionListData">Instance of favoriteDistributionListData.</param>
        /// <returns>Returns data model.</returns>
        public async Task CreateOrUpdateFavoriteDistributionListDataAsync(
           FavoriteDistributionListData favoriteDistributionListData)
        {
            FavoriteDistributionListTableEntity favoriteDistributionListDataEntity = new FavoriteDistributionListTableEntity()
            {
                UserPrincipalName = favoriteDistributionListData.UserPrincipalName.ToLower(),
                GroupId = favoriteDistributionListData.Id,
                PinStatus = favoriteDistributionListData.IsPinned,
            };

            await this.CreateOrUpdateAsync(favoriteDistributionListDataEntity);
        }

        /// <summary>
        /// Gets distribution list data from MS Graph based on search query.
        /// </summary>
        /// <param name="query">query to be used for filter.</param>
        /// <returns>DistributionList model data.</returns>
        public async Task<List<DistributionList>> GetDistributionListByQueryAsync(
            string query)
        {
            string[] scopes = { "Group.Read.All" };
            string graphQuery = $"https://graph.microsoft.com/v1.0/groups?$top=100&$select=id,displayName,mail,mailNickname,mailEnabled,proxyAddresses&$filter=startswith(displayName,'{query}')";
            List<DistributionList> distributionList = await this.protectedApiCallHelper.CallGraphApiOnBehalfOfUser<DistributionList>(scopes, graphQuery);

            return distributionList;
        }

        /// <summary>
        /// Gets distribution list details by group id from AAD.
        /// </summary>
        /// <param name="groupId">Distribution list group Id.</param>
        /// <returns>Distribution list model data.</returns>
        public async Task<DistributionList> GetADGroupDetailsAsync(
            string groupId)
        {
            string[] scopes = { "Group.Read.All" };
            string graphQuery = $"https://graph.microsoft.com/v1.0/groups/{groupId}?$select=id,displayName,mail,mailNickname,mailEnabled";
            return await this.protectedApiCallHelper.CallGraphApiOnBehalfOfUserWithDirectJToken<DistributionList>(scopes, graphQuery, true);
        }

        /// <summary>
        /// Get distribution list details and members count from AAD.
        /// </summary>
        /// <param name="groupIds">List of distribution list ids.</param>
        /// <returns>Count of members in distribution list.</returns>
        public async Task<List<DistributionList>> GetADGroupMemberCountBatchAsync(
            List<string> groupIds)
        {
            IEnumerable<List<string>> groupBatches = groupIds.SplitList(5);
            List<DistributionList> dlGraphData = new List<DistributionList>();
            string[] queries = new string[2] { "/groups/{0}/members?$select=displayName,mail,userType&$top=999", "/groups/{0}?$select=id,displayName,mail,mailNickname,mailEnabled" };

            foreach (List<string> groupBatch in groupBatches)
            {
                string[] scopes = { "Group.Read.All" };
                string graphQuery = $"https://graph.microsoft.com/v1.0/$batch";
                List<MSGraphBatchRequest> allRequests = MSGraphBatchRequestCreator.CreateBatchRequestPayloadForGroups(queries, "GET", groupBatch);
                MSGraphBatchRequestPayload payload = new MSGraphBatchRequestPayload()
                {
                    Requests = allRequests,
                };

                List<MSGraphBatchResponse<dynamic>> responses = await this.protectedApiCallHelper.CallGraphApiPostOnBehalfOfUser(scopes, graphQuery, JsonConvert.SerializeObject(payload));
                if (responses != null)
                {
                    foreach (string group in groupBatch)
                    {
                        try
                        {
                            List<DistributionListMember> dlMember = this.protectedApiCallHelper.GetValue<List<DistributionListMember>>(JsonConvert.SerializeObject(responses.Find(s => s.Id == group).Body), "value");
                            DistributionList dlDetails = this.protectedApiCallHelper.GetValue<DistributionList>(JsonConvert.SerializeObject(responses.Find(s => s.Id == group + '1')), "body");
                            if (dlDetails != null && dlMember != null)
                            {
                                dlDetails.NoOfMembers = dlMember.Where(x => (string.Equals(x.Type, "#microsoft.graph.group", StringComparison.OrdinalIgnoreCase) ||
                                                                             string.Equals(x.UserType, "member", StringComparison.OrdinalIgnoreCase))).Count();

                                dlGraphData.Add(dlDetails);
                            }
                        }
                        catch (Exception ex)
                        {
                            this.telemetryClient.TrackTrace($"An error occurred in GetADGroupMemberCountBatchAsync:  {group}", SeverityLevel.Error);
                            this.telemetryClient.TrackException(ex);
                        }
                    }
                }
            }

            return dlGraphData;
        }
    }
}
