// <copyright file="FavoriteDistributionListTableRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Repositories
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Configuration;

    /// <summary>
    /// Repository of the favorite distribution lists in the table storage.
    /// </summary>
    public class FavoriteDistributionListTableRepository : BaseRepository<FavoriteDistributionListTableEntity>
    {
        private static readonly string TableName = "FavoriteDistributionLists";

        /// <summary>
        /// Initializes a new instance of the <see cref="FavoriteDistributionListTableRepository"/> class.
        /// </summary>
        /// <param name="configuration">Singleton instance of application configuration.</param>
        public FavoriteDistributionListTableRepository(
            IConfiguration configuration)
            : base(
                configuration,
                TableName)
        {
        }

        /// <summary>
        /// To query favorite distribution lists from table storage based on user principal name.
        /// </summary>
        /// <param name="userPrincipalName">User Principal Name.</param>
        /// <returns>Favorite distribution lists of the user from database. </returns>
        public async Task<IEnumerable<FavoriteDistributionListTableEntity>> GetAllFavoriteDistributionListsAsync(string userPrincipalName)
        {
            IEnumerable<FavoriteDistributionListTableEntity> result = await this.GetAllAsync(userPrincipalName);
            return result;
        }
    }
}
