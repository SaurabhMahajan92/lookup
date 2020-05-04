﻿// <copyright file="FavoriteDistributionListMemberTableEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Repositories
{
    using Microsoft.Azure.Cosmos.Table;

    /// <summary>
    /// Favorite Distribution List Member Data table entity class represents pinned member records.
    /// </summary>
    public class FavoriteDistributionListMemberTableEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets pinned record's distribution list GUID.
        /// </summary>
        public string DistributionListID { get; set; }

        /// <summary>
        /// Gets or sets Partition key with users's user principal name.
        /// </summary>
        [IgnoreProperty]
        public string UserPrincipalName
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }

        /// <summary>
        /// Gets or sets row key with pinned record id + Distribution list id.
        /// </summary>
        [IgnoreProperty]
        public string DistributionListMemberId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }
    }
}