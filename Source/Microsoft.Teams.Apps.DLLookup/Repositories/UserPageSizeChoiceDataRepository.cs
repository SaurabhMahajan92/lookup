// <copyright file="UserPageSizeChoiceDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Repositories
{
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.DLLookup.Models;

    /// <summary>
    /// This class helps to read/write page size for currently logged in user.
    /// </summary>
    public class UserPageSizeChoiceDataRepository : UserPageSizeChoiceTableRepository
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UserPageSizeChoiceDataRepository"/> class.
        /// </summary>
        /// <param name="configuration">Singleton instance of application configuration.</param>
        public UserPageSizeChoiceDataRepository(
            IConfiguration configuration)
            : base(
                configuration)
        {
        }

        /// <summary>
        /// This method is used to store page size into database.
        /// </summary>
        /// <param name="userPrincipalName">User Principal Name of the logged in User.</param>
        /// <param name="pageSize">Page size to be stored.</param>
        /// <param name="pageType">Page for which the page size needs to be stored.</param>
        /// <returns>A <see cref="Task"/> representing the result of the asynchronous operation.</returns>
        public async Task CreateOrUpdateUserPageSizeChoiceDataAsync(
                string userPrincipalName,
                int pageSize,
                PageType pageType)
        {
            UserPageSizeChoiceTableEntity userPageSizeChoiceDataEntity = await this.GetAsync("default", userPrincipalName);
            if (userPageSizeChoiceDataEntity == null)
            {
                userPageSizeChoiceDataEntity = new UserPageSizeChoiceTableEntity();
            }

            userPageSizeChoiceDataEntity.DefaultValue = "default";
            userPageSizeChoiceDataEntity.UserPrincipalName = userPrincipalName.ToLower();

            if (pageType == PageType.DistributionList)
            {
                userPageSizeChoiceDataEntity.DistributionListPageSize = pageSize;
            }
            else
            {
                userPageSizeChoiceDataEntity.DistributionListMemberPageSize = pageSize;
            }

            await this.CreateOrUpdateAsync(userPageSizeChoiceDataEntity);
        }
    }
}
