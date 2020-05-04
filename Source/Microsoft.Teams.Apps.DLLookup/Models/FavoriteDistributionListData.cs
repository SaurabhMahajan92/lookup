﻿// <copyright file="FavoriteDistributionListData.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Models
{
    /// <summary>
    /// This model represents Favorite distribution lists.
    /// </summary>
    public class FavoriteDistributionListData
    {
        /// <summary>
        /// Gets or sets Id of the distribution lists in the favorites list.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether record is pinned or not by the logged in user.
        /// </summary>
        public bool IsPinned { get; set; }

        /// <summary>
        /// Gets or sets display name of the distribution lists in the favorites list.
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Gets or sets alias of the distribution lists in the favorites list.
        /// </summary>
        public string Mail { get; set; }

        /// <summary>
        /// Gets or sets number of contacts of the distribution lists in the favorites list.
        /// </summary>
        public int NoOfContacts { get; set; }

        /// <summary>
        /// Gets or sets user principal name of the distribution lists in the favorites list.
        /// </summary>
        public string UserPrincipalName { get; set; }
    }
}
