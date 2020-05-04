// <copyright file="FavoriteDistributionListTableEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Repositories
{
    using Microsoft.Azure.Cosmos.Table;

    /// <summary>
    /// Favorite Distribution List data table entity class used to represent pinned distribution list records.
    /// </summary>
    public class FavoriteDistributionListTableEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets a value indicating whether record is pinned or not.
        /// </summary>
        public bool PinStatus { get; set; }

        /// <summary>
        /// Gets or sets Row key with distribution list id.
        /// </summary>
        [IgnoreProperty]
        public string GroupId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets Partition key with user's user principal name.
        /// </summary>
        [IgnoreProperty]
        public string UserPrincipalName
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }
    }
}