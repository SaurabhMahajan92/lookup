// <copyright file="IPresenceDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Repositories
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.DLLookup.Models;

    /// <summary>
    /// Interface <see cref="IPresenceDataRepository"/> helps fetching user presence and contact data.
    /// </summary>
    public interface IPresenceDataRepository
    {
        /// <summary>
        /// Get User presence details in a batch.
        /// </summary>
        /// <param name="peoplePresenceDataArray">Array of People Presence Data object used to get presence information.</param>
        /// <returns>People Presence Data model data filled with presence information.</returns>
        Task<List<PeoplePresenceData>> GetBatchUserPresenceAsync(PeoplePresenceData[] peoplePresenceDataArray);

        /// <summary>
        /// Gets Online members count in a distribution list.
        /// </summary>
        /// <param name="groupId">Distribution list id.</param>
        /// <returns><see cref="Task{TResult}"/>Online members count in distribution list.</returns>
        Task<int> GetDistributionListMembersOnlineCountAsync(string groupId);
    }
}