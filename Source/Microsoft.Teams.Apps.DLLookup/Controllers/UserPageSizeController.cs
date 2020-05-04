// <copyright file="UserPageSizeController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Controllers
{
    using System;
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
    /// creating <see cref="UserPageSizeController"/> class with ControllerBase as base class. Controller for page size APIs.
    /// </summary>
    [Authorize]
    [Route("api/UserPageSize")]
    [ApiController]
    public class UserPageSizeController : ControllerBase
    {
        private readonly TelemetryClient telemetryClient;
        private readonly UserPageSizeChoiceDataRepository userPageSizeChoiceDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserPageSizeController"/> class.
        /// </summary>
        /// <param name="telemetryClient">Singleton TelemetryClient instance used to send telemetry to Azure application insights.</param>
        /// <param name="userPageSizeChoiceDataRepository">Singleton UserPageSizeChoiceDataRepository instance used to perform read/store operations for page size.</param>
        public UserPageSizeController(
            TelemetryClient telemetryClient,
            UserPageSizeChoiceDataRepository userPageSizeChoiceDataRepository)
        {
            this.telemetryClient = telemetryClient;
            this.userPageSizeChoiceDataRepository = userPageSizeChoiceDataRepository;
        }

        /// <summary>
        /// Gets the page size values for currently logged in user from database.
        /// </summary>
        /// <returns>A <see cref="Task"/> representing user page size.</returns>
        [HttpGet]
        public async Task<UserPageSizeChoiceTableEntity> GetUserPageSizeChoiceAsync()
        {
            UserPageSizeChoiceTableEntity userPageSizeChoiceDataEntity = new UserPageSizeChoiceTableEntity();
            try
            {
                string userPrincipalName = this.HttpContext.User.FindFirst(ClaimTypes.Upn)?.Value.ToLower();
                if (string.IsNullOrEmpty(userPrincipalName))
                {
                    this.telemetryClient.TrackTrace($"There's no user principal name claim.", SeverityLevel.Error);
                    this.HttpContext.Response.ContentType = "text/plain";
                    this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                    await this.HttpContext.Response.WriteAsync("There's no user principal name claim");
                }

                userPageSizeChoiceDataEntity = await this.userPageSizeChoiceDataRepository.GetUserPageSizeChoice(userPrincipalName);
                this.telemetryClient.TrackEvent($"Retrieved user Page size : {JsonConvert.SerializeObject(userPageSizeChoiceDataEntity)}");
            }
            catch (MsalException ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in getUserPageSizeChoice: {ex.Message}. Property: {this.HttpContext.User.FindFirst(ClaimTypes.Upn)?.Value.ToLower()}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                await this.HttpContext.Response.WriteAsync("An authentication error occurred while acquiring a token for downstream API\n" + ex.ErrorCode + "\n" + ex.Message);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in getUserPageSizeChoice: {ex.Message}. Property: {this.HttpContext.User.FindFirst(ClaimTypes.Upn)?.Value.ToLower()}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                await this.HttpContext.Response.WriteAsync("An error occurred while calling the downstream API\n" + ex.Message);
            }

            return userPageSizeChoiceDataEntity;
        }

        /// <summary>
        /// Stores page size values in database for currently logged in user.
        /// </summary>
        /// <param name="userPageSizeChoice">Page size to be stored.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [HttpPost]
        public async Task CreateUserPageSizeChoice([FromBody]UserPageSizeChoice userPageSizeChoice)
        {
            try
            {
                string userPrincipalName = this.HttpContext.User.FindFirst(ClaimTypes.Upn)?.Value.ToLower();
                if (string.IsNullOrEmpty(userPrincipalName))
                {
                    this.telemetryClient.TrackTrace($"There's no user principal name claim. Parameters:{userPageSizeChoice}", SeverityLevel.Error);
                    this.HttpContext.Response.ContentType = "text/plain";
                    this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                    await this.HttpContext.Response.WriteAsync("There's no user principal name claim");
                }

                await this.userPageSizeChoiceDataRepository.CreateOrUpdateUserPageSizeChoiceDataAsync(userPrincipalName, userPageSizeChoice.PageSize, userPageSizeChoice.PageId);

                this.telemetryClient.TrackEvent($"Created/Updated user Page size : {JsonConvert.SerializeObject(userPageSizeChoice)}");

                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.OK;
                await this.HttpContext.Response.WriteAsync("Choice successfully saved.");
            }
            catch (MsalException ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in createUserPageSizeChoice: {ex.Message}. Parameters: {this.HttpContext.User.FindFirst(ClaimTypes.Upn)?.Value.ToLower()}:{JsonConvert.SerializeObject(userPageSizeChoice)}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in createUserPageSizeChoice: {ex.Message}. Parameters:{this.HttpContext.User.FindFirst(ClaimTypes.Upn)?.Value.ToLower()}:{JsonConvert.SerializeObject(userPageSizeChoice)}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                await this.HttpContext.Response.WriteAsync("An error occurred while calling the downstream API\n" + ex.Message);
            }
        }
    }
}