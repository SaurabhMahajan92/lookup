// <copyright file="DistributionListMembersController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Http;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Identity.Client;
    using Microsoft.Teams.Apps.DLLookup.Models;
    using Microsoft.Teams.Apps.DLLookup.Repositories;
    using Newtonsoft.Json;

    /// <summary>
    /// Creating <see cref="DistributionListMembersController"/> class with ControllerBase as base class. Controller for Distribution List member APIs.
    /// </summary>
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class DistributionListMembersController : ControllerBase
    {
        private readonly TelemetryClient telemetryClient;
        private readonly FavoriteDistributionListMemberDataRepository favoriteDistributionListMemberDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="DistributionListMembersController"/> class.
        /// </summary>
        /// <param name="telemetryClient">Singleton TelemetryClient instance used to send telemetry to Azure application insights.</param>
        /// <param name="favoriteDistributionListMemberDataRepository">Scoped FavoriteDistributionListMemberDataRepository instance used to read/write distribution list member related operations.</param>
        public DistributionListMembersController(
            TelemetryClient telemetryClient,
            FavoriteDistributionListMemberDataRepository favoriteDistributionListMemberDataRepository)
        {
            this.telemetryClient = telemetryClient;
            this.favoriteDistributionListMemberDataRepository = favoriteDistributionListMemberDataRepository;
        }

        /// <summary>
        /// Gets the members in a distribution list using the group GUID from Graph API.
        /// </summary>
        /// <param name="groupId">Distribution list group GUID.</param>
        /// <returns><DistributionListMember>A <see cref="Task"/> list of distribution list members information.</DistributionListMember></returns>
        [HttpGet]
        public async Task<List<DistributionListMember>> GetDistributionListMembersDataAsync([FromQuery]string groupId)
        {
            List<DistributionListMember> distributionListMembers = new List<DistributionListMember>();
            try
            {
                string userPrincipalName = this.HttpContext.User.FindFirst(ClaimTypes.Upn)?.Value.ToLower();
                if (string.IsNullOrEmpty(userPrincipalName))
                {
                    this.telemetryClient.TrackTrace($"There's no user principal name claim. Parameters:{groupId}", SeverityLevel.Error);
                    this.HttpContext.Response.ContentType = "text/plain";
                    this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                    await this.HttpContext.Response.WriteAsync("There's no user principal name claim.");
                }

                distributionListMembers = await this.favoriteDistributionListMemberDataRepository.GetDistributionListMembersDataAsync(
                    groupId,
                    userPrincipalName);
            }
            catch (MsalException ex)
            {
                this.telemetryClient.TrackTrace($"A Msal error occurred in GetDistributionListMembersDataAsync: {ex.Message}, Parameters:{groupId}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                await this.HttpContext.Response.WriteAsync("An authentication error occurred while acquiring a token for downstream API\n" + ex.ErrorCode + "\n" + ex.Message);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in GetDistributionListMembersDataAsync: {ex.Message}, Parameters:{groupId}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                await this.HttpContext.Response.WriteAsync("An error occurred while calling the downstream API\n" + ex.Message);
            }

            return distributionListMembers;
        }

        /// <summary>
        /// Adds member data to the table storage on being pinned by the user.
        /// </summary>
        /// <param name="favoriteDistributionListMemberData">Instance of favorite distribution list member data holding the values sent by the user.</param>
        /// <returns><Task>A <see cref="Task"/> representing the asynchronous operation.</Task></returns>
        [HttpPost]
        public async Task CreateFavoriteDistributionMemberListData([FromBody]FavoriteDistributionListMemberData favoriteDistributionListMemberData)
        {
            try
            {
                string userPrincipalName = this.HttpContext.User.FindFirst(ClaimTypes.Upn)?.Value.ToLower();

                if (string.IsNullOrEmpty(userPrincipalName))
                {
                    this.telemetryClient.TrackTrace($"There's no user principal name claim.", SeverityLevel.Error);
                    this.HttpContext.Response.ContentType = "text/plain";
                    this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                    await this.HttpContext.Response.WriteAsync("There's no user principal name claim.");
                }

                favoriteDistributionListMemberData.UserPrincipalName = userPrincipalName;

                await this.favoriteDistributionListMemberDataRepository.CreateOrUpdateFavoriteDistributionListMemberDataAsync(favoriteDistributionListMemberData);

                this.telemetryClient.TrackEvent($"Updated PIN status : {JsonConvert.SerializeObject(favoriteDistributionListMemberData)}");

                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.OK;
                await this.HttpContext.Response.WriteAsync("Pinned user successfully.");
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in CreateFavoriteDistributionMemberListData: {ex.Message}, Parameters: {JsonConvert.SerializeObject(favoriteDistributionListMemberData)}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                await this.HttpContext.Response.WriteAsync("An error occurred while calling the downstream API\n" + ex.Message);
            }
        }

        /// <summary>
        /// Updates azure table storage when user unpins their favorite members.
        /// </summary>
        /// <param name="favoriteDistributionListMemberData">Instance of FavoriteDistributionListMemberData holding the values sent by the user for unpin.</param>
        /// <returns><Task>A <see cref="Task"/> representing the asynchronous operation.</Task></returns>
        [HttpDelete]
        public async Task DeleteFavoriteDistributionListMemberDataAsync([FromBody]FavoriteDistributionListMemberData favoriteDistributionListMemberData)
        {
            try
            {
                string userPrincipalName = this.HttpContext.User.FindFirst(ClaimTypes.Upn)?.Value.ToLower();

                if (string.IsNullOrEmpty(userPrincipalName))
                {
                    this.telemetryClient.TrackTrace($"There's no user principal name claim.", SeverityLevel.Error);
                    this.HttpContext.Response.ContentType = "text/plain";
                    this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                    await this.HttpContext.Response.WriteAsync("There's no user principal name claim.");
                }

                favoriteDistributionListMemberData.UserPrincipalName = userPrincipalName;

                FavoriteDistributionListMemberTableEntity favoriteDistributionListMemberDataEntity = await this.favoriteDistributionListMemberDataRepository.GetAsync(
                   favoriteDistributionListMemberData.UserPrincipalName.ToLower(),
                   favoriteDistributionListMemberData.PinnedUserId + favoriteDistributionListMemberData.DistributionListID);

                if (favoriteDistributionListMemberDataEntity != null)
                {
                    await this.favoriteDistributionListMemberDataRepository.DeleteAsync(favoriteDistributionListMemberDataEntity);
                    this.telemetryClient.TrackEvent($"Deleted favorite user : {JsonConvert.SerializeObject(favoriteDistributionListMemberData)}");
                }
                else
                {
                    this.telemetryClient.TrackEvent($"Did not find favorite user to delete : {JsonConvert.SerializeObject(favoriteDistributionListMemberData)}");
                }

                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.OK;
                await this.HttpContext.Response.WriteAsync("Unpinned user successfully.");
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in DeleteFavoriteDistributionListMemberDataAsync: {ex.Message}, Parameters:{JsonConvert.SerializeObject(favoriteDistributionListMemberData)}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                await this.HttpContext.Response.WriteAsync("An error occurred while calling the downstream API\n" + ex.Message);
            }
        }
    }
}