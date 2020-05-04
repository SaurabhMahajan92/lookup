// <copyright file="DistributionListsController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
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
    /// Creating <see cref="DistributionListsController"/> class with ControllerBase as base class. Controller for Distribution List APIs.
    /// </summary>
    [Authorize]
    [Route("api/distributionLists")]
    [ApiController]
    public class DistributionListsController : ControllerBase
    {
        private readonly TelemetryClient telemetryClient;
        private readonly FavoriteDistributionListDataRepository favoriteDistributionListDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="DistributionListsController"/> class.
        /// </summary>
        /// <param name="telemetryClient">Singleton TelemetryClient instance used to send telemetry to Azure application insights.</param>
        /// <param name="favoriteDistributionListDataRepository">Scoped favoriteDistributionListDataRepository instance used to read/write distribution list related operations.</param>
        public DistributionListsController(
            TelemetryClient telemetryClient,
            FavoriteDistributionListDataRepository favoriteDistributionListDataRepository)
        {
            this.telemetryClient = telemetryClient;
            this.favoriteDistributionListDataRepository = favoriteDistributionListDataRepository;
        }

        /// <summary>
        /// Gets distribution lists from AAD by search query.
        /// </summary>
        /// <param name="query">Search query.</param>
        /// <returns>A <see cref="Task"/>List of distribution lists information.</returns>
        [HttpGet]
        [Route("getDistributionList")]
        public async Task<List<DistributionList>> GetDistributionListByQueryAsync([FromQuery]string query)
        {
            List<DistributionList> distributionList = new List<DistributionList>();
            try
            {
                return await this.favoriteDistributionListDataRepository.GetDistributionListByQueryAsync(query);
            }
            catch (MsalException ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in GetDistributionListByQueryAsync: {ex.Message}. Parameters:{query}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                await this.HttpContext.Response.WriteAsync("An authentication error occurred while acquiring a token for downstream API\n" + ex.ErrorCode + "\n" + ex.Message);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in GetDistributionListByQueryAsync: {ex.Message}. Parameters:{query}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                await this.HttpContext.Response.WriteAsync("An error occurred while calling the downstream API\n" + ex.Message);
            }

            return distributionList;
        }

        /// <summary>
        /// Gets all favorite distribution lists which includes data from database and Graph API.
        /// </summary>
        /// <returns>A <see cref="Task"/> list of favorite distribution lists information.</returns>
        [HttpGet]
        public async Task<List<FavoriteDistributionListData>> GetAllFavoriteDistributionListDataAsync()
        {
            List<FavoriteDistributionListData> favoriteDistributionListData = new List<FavoriteDistributionListData>();
            List<DistributionList> output = new List<DistributionList>();
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

                IEnumerable<FavoriteDistributionListTableEntity> favoriteDistributionListDataEntities = await this.favoriteDistributionListDataRepository.GetAllFavoriteDistributionListsAsync(userPrincipalName);
                if (favoriteDistributionListDataEntities.Count() > 0)
                {
                    List<string> groupIds = favoriteDistributionListDataEntities.ToList().Select(dl => dl.GroupId).ToList();
                    output = await this.favoriteDistributionListDataRepository.GetADGroupMemberCountBatchAsync(groupIds);

                    foreach (FavoriteDistributionListTableEntity currentItem in favoriteDistributionListDataEntities)
                    {
                        try
                        {
                            DistributionList currentDistributionList = output.Find(dl => dl.Id == currentItem.GroupId);
                            favoriteDistributionListData.Add(
                                new FavoriteDistributionListData
                                {
                                    UserPrincipalName = currentItem.UserPrincipalName.ToLower(),
                                    IsPinned = currentItem.PinStatus,
                                    DisplayName = currentDistributionList.DisplayName,
                                    Mail = currentDistributionList.Mail,
                                    NoOfContacts = currentDistributionList.NoOfMembers,
                                    Id = currentItem.GroupId,
                                });
                        }
                        catch (Exception)
                        {
                            // Exception occurred as resource has not been found in AD anymore. Hence skip it.
                            this.telemetryClient.TrackTrace($"Resource has not been found in AD anymore. Hence skipping it. {JsonConvert.SerializeObject(currentItem)}", SeverityLevel.Error);
                        }
                    }

                    return favoriteDistributionListData;
                }
            }
            catch (MsalException ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in GetAllFavoriteDistributionListDataAsync: {ex.Message}. Property: {this.HttpContext.User.FindFirst(ClaimTypes.Upn)?.Value.ToLower()}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                await this.HttpContext.Response.WriteAsync("An authentication error occurred while acquiring a token for downstream API\n" + ex.ErrorCode + "\n" + ex.Message);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in GetAllFavoriteDistributionListDataAsync: {ex.Message}. Property: {this.HttpContext.User.FindFirst(ClaimTypes.Upn)?.Value.ToLower()}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                await this.HttpContext.Response.WriteAsync("An error occurred while calling the downstream API\n" + ex.Message);
            }

            return favoriteDistributionListData;
        }

        /// <summary>
        /// Adds favorite distribution lists to database as user favorites.
        /// </summary>
        /// <param name="distributionListDetails">Distribution list array to be saved as user favorite.</param>
        /// <returns>>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [HttpPost]
        public async Task CreateFavoriteDistributionListDataAsync([FromBody]FavoriteDistributionListData[] distributionListDetails)
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

                for (int i = 0; i < distributionListDetails.Length; i++)
                {
                    distributionListDetails[i].UserPrincipalName = userPrincipalName;
                }

                bool isErrorOccurred = false;

                foreach (FavoriteDistributionListData currentItem in distributionListDetails)
                {
                    try
                    {
                        await this.favoriteDistributionListDataRepository.CreateOrUpdateFavoriteDistributionListDataAsync(currentItem);
                    }
                    catch (Exception ex)
                    {
                        isErrorOccurred = true;
                        this.telemetryClient.TrackTrace($"An error occurred in CreateFavoriteDistributionListDataAsync: {ex.Message}. Parameters: {JsonConvert.SerializeObject(currentItem)}", SeverityLevel.Error);
                        this.telemetryClient.TrackException(ex);
                    }
                }

                if (!isErrorOccurred)
                {
                    this.telemetryClient.TrackEvent($"Added favorite DLs : {JsonConvert.SerializeObject(distributionListDetails)}");
                    this.HttpContext.Response.ContentType = "text/plain";
                    this.HttpContext.Response.StatusCode = (int)HttpStatusCode.OK;
                    await this.HttpContext.Response.WriteAsync("Added new favorite distribution lists.");
                }
                else
                {
                    await this.HttpContext.Response.WriteAsync($"Error occurred in adding new favorite distribution lists. {JsonConvert.SerializeObject(distributionListDetails)}");
                }
            }
            catch (MsalException ex)
            {
                this.telemetryClient.TrackTrace($"A msal error occurred in CreateFavoriteDistributionListDataAsync: {ex.Message}. Parameters: {JsonConvert.SerializeObject(distributionListDetails)}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                await this.HttpContext.Response.WriteAsync("An authentication error occurred while acquiring a token for downstream API\n" + ex.ErrorCode + "\n" + ex.Message);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in CreateFavoriteDistributionListDataAsync: {ex.Message}. Parameters: {JsonConvert.SerializeObject(distributionListDetails)}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                await this.HttpContext.Response.WriteAsync("An error occurred while calling the downstream API\n" + ex.Message);
            }
        }

        /// <summary>
        /// Updates distribution list favorite status (Pin/Unpin) in database.
        /// </summary>
        /// <param name="favoriteDistributionListData">Distribution list data used to update pin status for currently logged in user.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [HttpPut]
        public async Task UpdateFavoriteDistributionListData([FromBody]FavoriteDistributionListData favoriteDistributionListData)
        {
            try
            {
                string userPrincipalName = this.HttpContext.User.FindFirst(ClaimTypes.Upn)?.Value.ToLower();

                if (string.IsNullOrEmpty(userPrincipalName))
                {
                    this.telemetryClient.TrackTrace("There's no user principal name claim.", SeverityLevel.Error);
                    this.HttpContext.Response.ContentType = "text/plain";
                    this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                    await this.HttpContext.Response.WriteAsync("There's no user principal name claim.");
                }

                favoriteDistributionListData.UserPrincipalName = userPrincipalName;

                await this.favoriteDistributionListDataRepository.CreateOrUpdateFavoriteDistributionListDataAsync(favoriteDistributionListData);

                this.telemetryClient.TrackEvent($"Updated pin status for DL : {JsonConvert.SerializeObject(favoriteDistributionListData)}");

                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.OK;
                await this.HttpContext.Response.WriteAsync("Distribution List pin status updated.");
            }
            catch (MsalException ex)
            {
                this.telemetryClient.TrackTrace($"A msal error occurred in UpdateFavoriteDistributionListData: {ex.Message}. Parameters: {JsonConvert.SerializeObject(favoriteDistributionListData)}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                await this.HttpContext.Response.WriteAsync("An authentication error occurred while acquiring a token for downstream API\n" + ex.ErrorCode + "\n" + ex.Message);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in UpdateFavoriteDistributionListData: {ex.Message}. Parameters: {JsonConvert.SerializeObject(favoriteDistributionListData)}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                await this.HttpContext.Response.WriteAsync("An error occurred while calling the downstream API\n" + ex.Message);
            }
        }

        /// <summary>
        /// Deletes favorite distribution list from database.
        /// </summary>
        /// <param name="favoriteDistributionListData">Distribution list data to delete.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [HttpDelete]
        public async Task DeleteFavoriteDistributionListDataAsync([FromBody]FavoriteDistributionListData favoriteDistributionListData)
        {
            try
            {
                string userPrincipleName = this.HttpContext.User.FindFirst(ClaimTypes.Upn)?.Value.ToLower();

                if (string.IsNullOrEmpty(userPrincipleName))
                {
                    this.telemetryClient.TrackTrace($"There's no user principal name claim.", SeverityLevel.Error);
                    this.HttpContext.Response.ContentType = "text/plain";
                    this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                    await this.HttpContext.Response.WriteAsync("There's no user principal name claim.");
                }

                favoriteDistributionListData.UserPrincipalName = userPrincipleName;

                FavoriteDistributionListTableEntity favoriteDistributionListDataEntity = await this.favoriteDistributionListDataRepository.GetAsync(
                   userPrincipleName,
                   favoriteDistributionListData.Id);
                if (favoriteDistributionListDataEntity != null)
                {
                    await this.favoriteDistributionListDataRepository.DeleteAsync(favoriteDistributionListDataEntity);
                    this.telemetryClient.TrackEvent($"Deleted favorite DL : {JsonConvert.SerializeObject(favoriteDistributionListData)}");
                }
                else
                {
                    this.telemetryClient.TrackEvent($"Did not find favorite user to delete : {JsonConvert.SerializeObject(favoriteDistributionListData)}");
                }

                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.OK;
                await this.HttpContext.Response.WriteAsync("Deleted Favorite distribution list successfully.");
            }
            catch (MsalException ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in DeleteFavoriteDistributionListDataAsync: {ex.Message}. Parameters:{JsonConvert.SerializeObject(favoriteDistributionListData)}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                await this.HttpContext.Response.WriteAsync("An authentication error occurred while acquiring a token for downstream API\n" + ex.ErrorCode + "\n" + ex.Message);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in DeleteFavoriteDistributionListDataAsync: {ex.Message}. Parameters:{JsonConvert.SerializeObject(favoriteDistributionListData)}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                await this.HttpContext.Response.WriteAsync("An error occurred while calling the downstream API\n" + ex.Message);
            }
        }
    }
}