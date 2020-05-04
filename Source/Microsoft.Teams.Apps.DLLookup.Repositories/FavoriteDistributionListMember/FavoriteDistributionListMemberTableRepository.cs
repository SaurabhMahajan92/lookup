// <copyright file="FavoriteDistributionListMemberTableRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Repositories
{
    using Microsoft.Extensions.Configuration;

    /// <summary>
    /// Repository of the favorite distribution list members in the table storage.
    /// </summary>
    public class FavoriteDistributionListMemberTableRepository : BaseRepository<FavoriteDistributionListMemberTableEntity>
    {
        private static readonly string TableName = "FavoriteDistributionListMembers";

        /// <summary>
        /// Initializes a new instance of the <see cref="FavoriteDistributionListMemberTableRepository"/> class.
        /// </summary>
        /// <param name="configuration">Singleton instance of application configuration.</param>
        public FavoriteDistributionListMemberTableRepository(
            IConfiguration configuration)
            : base(
                configuration,
                TableName)
        {
        }
    }
}
