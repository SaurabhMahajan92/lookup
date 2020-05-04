// <copyright file="FavoriteDistributionListMemberDataRepository.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Repositories
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.DLLookup.Models;

    /// <summary>
    /// This class contains read/write operations for distribution list member on AAD and table storage.
    /// </summary>
    public class FavoriteDistributionListMemberDataRepository : FavoriteDistributionListMemberTableRepository
    {
        private readonly IProtectedApiCallHelper protectedApiCallHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="FavoriteDistributionListMemberDataRepository"/> class.
        /// </summary>
        /// <param name="configuration">Singleton instance of application configuration.</param>
        /// <param name="protectedApiCallHelper">Scoped protectedApiCallHelper instance used to call Web APIs.</param>
        public FavoriteDistributionListMemberDataRepository(
            IConfiguration configuration, IProtectedApiCallHelper protectedApiCallHelper)
            : base(
                configuration)
        {
            this.protectedApiCallHelper = protectedApiCallHelper;
        }

        /// <summary>
        /// Create/Update favorite distribution list member data to table storage.
        /// </summary>
        /// <param name="favoriteDistributionListMemberData">Favorite distribution list member data to be stored in database.</param>
        /// <returns>A <see cref="Task"/> representing the result of the asynchronous operation.</returns>
        public async Task CreateOrUpdateFavoriteDistributionListMemberDataAsync(
           FavoriteDistributionListMemberData favoriteDistributionListMemberData)
        {
            var favoriteDistributionListMemberDataEntity = new FavoriteDistributionListMemberTableEntity()
            {
                UserPrincipalName = favoriteDistributionListMemberData.UserPrincipalName.ToLower(),
                DistributionListMemberId = (favoriteDistributionListMemberData.PinnedUserId + favoriteDistributionListMemberData.DistributionListID).ToLower(),
                DistributionListID = favoriteDistributionListMemberData.DistributionListID,
            };

            await this.CreateOrUpdateAsync(favoriteDistributionListMemberDataEntity);
        }

        /// <summary>
        /// Gets distribution list members data from AAD and database.
        /// </summary>
        /// <param name="groupID">distribution list id to filter records.</param>
        /// <param name="userPrincipalName">User Principal Name to filter pinned records from database.</param>
        /// <returns>returns DistributionListMember model data containing members and users favorite information.</returns>
        public async Task<List<DistributionListMember>> GetDistributionListMembersDataAsync(
            string groupID,
            string userPrincipalName)
        {
            string[] scopes = { "Group.Read.All" };
            string graphQuery = $"https://graph.microsoft.com/v1.0/groups/{groupID}//members?$top=100&$select=id,displayName,jobTitle,mail,userPrincipalName,userType";
            List<DistributionListMember> distributionListMemberList = await this.protectedApiCallHelper.CallGraphApiOnBehalfOfUser<DistributionListMember>(scopes, graphQuery);

            IEnumerable<FavoriteDistributionListMemberTableEntity> favoriteDistributionListMemberEntity = await this.GetAllAsync(userPrincipalName.ToLower());
            foreach (DistributionListMember member in distributionListMemberList)
            {
                string distributionListMemberId = member.Id + groupID;
                foreach (FavoriteDistributionListMemberTableEntity entity in favoriteDistributionListMemberEntity)
                {
                    if (entity.DistributionListMemberId == distributionListMemberId)
                    {
                        member.IsPinned = true;
                    }
                }
            }

            return distributionListMemberList.Where(x => x.Type == "#microsoft.graph.group" || string.Equals(x.UserType, "member", StringComparison.OrdinalIgnoreCase)).ToList();
        }
    }
}
