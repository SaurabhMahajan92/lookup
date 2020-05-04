// <copyright file="FavoriteDistributionListMemberData.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Models
{
    /// <summary>
    /// This model represents favorite distribution list member data.
    /// </summary>
    public class FavoriteDistributionListMemberData
    {
        /// <summary>
        /// Gets or sets user principal name of the favorite member in the distributed list.
        /// </summary>
        public string UserPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets user id of the favorite member in the distributed list.
        /// </summary>
        public string PinnedUserId { get; set; }

        /// <summary>
        /// Gets or sets distribution list GUID, the pinned member belongs to.
        /// </summary>
        public string DistributionListID { get; set; }
    }
}
