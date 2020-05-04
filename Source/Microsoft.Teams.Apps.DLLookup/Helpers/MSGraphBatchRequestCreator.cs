// <copyright file="MSGraphBatchRequestCreator.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Helpers
{
    using System.Collections.Generic;
    using Microsoft.Teams.Apps.DLLookup.Models;

    /// <summary>
    /// Class used for creating batch of requests used for getting information using MS Graph.
    /// </summary>
    public static class MSGraphBatchRequestCreator
    {
        /// <summary>
        /// This method is to create batch request for Graph API calls.
        /// </summary>
        /// <param name="queries">List of graph queries.</param>
        /// <param name="httpMethod">String denoting HTTP method.</param>
        /// <param name="groupIds">List of strings of group Ids.</param>
        /// <returns>A <see cref="MSGraphBatchRequest"/> representing the batch request for Graph API calls.</returns>
        public static List<MSGraphBatchRequest> CreateBatchRequestPayloadForGroups(string[] queries, string httpMethod, List<string> groupIds)
        {
            List<MSGraphBatchRequest> request = new List<MSGraphBatchRequest>();
            int queryCount = 0;
            foreach (string query in queries)
            {
                foreach (string groupId in groupIds)
                {
                    string id = groupId;
                    if (queryCount > 0)
                    {
                        id = groupId + queryCount;
                    }

                    MSGraphBatchRequest batchRequest = new MSGraphBatchRequest()
                    {
                        Id = id,
                        Method = httpMethod,
                        URL = string.Format(query, groupId),
                    };
                    request.Add(batchRequest);
                }

                queryCount++;
            }

            return request;
        }

        /// <summary>
        /// This method is to create batch request for Graph API calls.
        /// </summary>
        /// <param name="query">Graph query.</param>
        /// <param name="httpMethod">String denoting HTTP method.</param>
        /// <param name="peoplePresenceDataList">List of strings of user principal names.</param>
        /// <returns>A <see cref="MSGraphBatchRequest"/> representing the batch request for Graph API calls.</returns>
        public static List<MSGraphBatchRequest> CreateBatchRequestPayloadForPresence(string query, string httpMethod, List<PeoplePresenceData> peoplePresenceDataList)
        {
            List<MSGraphBatchRequest> request = new List<MSGraphBatchRequest>();
            foreach (PeoplePresenceData peoplePresenceData in peoplePresenceDataList)
            {
                MSGraphBatchRequest batchRequest = new MSGraphBatchRequest()
                {
                    Id = peoplePresenceData.UserPrincipalName,
                    Method = httpMethod,
                    URL = string.Format(query, peoplePresenceData.Id),
                };
                request.Add(batchRequest);
            }

            return request;
        }
    }
}
