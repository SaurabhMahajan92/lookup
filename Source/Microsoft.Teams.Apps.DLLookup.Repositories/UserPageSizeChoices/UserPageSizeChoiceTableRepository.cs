// <copyright file="UserPageSizeChoiceTableRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Repositories
{
    using System.Threading.Tasks;
    using Microsoft.Extensions.Configuration;

    /// <summary>
    /// Repository of the user page size choice in the table storage.
    /// </summary>
    public class UserPageSizeChoiceTableRepository : BaseRepository<UserPageSizeChoiceTableEntity>
    {
        private static readonly string TableName = "UserPageSizeChoices";

        /// <summary>
        /// Initializes a new instance of the <see cref="UserPageSizeChoiceTableRepository"/> class.
        /// </summary>
        /// <param name="configuration">Singleton instance of application configuration.</param>
        public UserPageSizeChoiceTableRepository(
            IConfiguration configuration)
            : base(
                configuration,
                TableName)
        {
        }

        /// <summary>
        /// To query page size information of a particular user from table storage.
        /// </summary>
        /// <param name="userPrincipalName">user's user principal name to query from database.</param>
        /// <returns>Distribution list and distribution list members page size.</returns>
        public async Task<UserPageSizeChoiceTableEntity> GetUserPageSizeChoice(string userPrincipalName)
        {
            UserPageSizeChoiceTableEntity result = await this.GetAsync("default", userPrincipalName);
            return result;
        }
    }
}
