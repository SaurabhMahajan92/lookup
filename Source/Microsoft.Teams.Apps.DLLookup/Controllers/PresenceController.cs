// <copyright file="PresenceController.cs" company="Microsoft">
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
    /// creating <see cref="PresenceController"/> class with ControllerBase as base class. Controller for user presence APIs.
    /// </summary>
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class PresenceController : ControllerBase
    {
        private readonly TelemetryClient telemetryClient;
        private readonly IPresenceDataRepository presenceDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="PresenceController"/> class.
        /// </summary>
        /// <param name="telemetryClient">Singleton TelemetryClient instance used to send telemetry to Azure application insights.</param>
        /// <param name="presenceDataRepository">Scoped PresenceDataRepository instance used to get presence information.</param>
        public PresenceController(
            TelemetryClient telemetryClient,
            IPresenceDataRepository presenceDataRepository)
        {
            this.telemetryClient = telemetryClient;
            this.presenceDataRepository = presenceDataRepository;
        }

        /// <summary>
        /// Get User presence status details.
        /// </summary>
        /// <param name="peoplePresenceDataArray">Array of People Presence Data object used to get presence information.</param>
        /// <returns>People Presence Data model data filled with presence information.</returns>
        [HttpPost]
        [Route("GetUserPresence")]
        public async Task<List<PeoplePresenceData>> GetUserPresenceAsync([FromBody]PeoplePresenceData[] peoplePresenceDataArray)
        {
            try
            {
                List<PeoplePresenceData> peoplePresenceDataList = await this.presenceDataRepository.GetBatchUserPresenceAsync(peoplePresenceDataArray);
                return peoplePresenceDataList;
            }
            catch (MsalException ex)
            {
                this.telemetryClient.TrackTrace($"A Msal error occurred in GetUserPresenceAsync: {ex.Message}, Parameters: {JsonConvert.SerializeObject(peoplePresenceDataArray)}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                await this.HttpContext.Response.WriteAsync("An authentication error occurred while acquiring a token for downstream API\n" + ex.ErrorCode + "\n" + ex.Message);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in GetUserPresenceAsync: {ex.Message}, Parameters: {JsonConvert.SerializeObject(peoplePresenceDataArray)}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                await this.HttpContext.Response.WriteAsync("An error occurred while calling the downstream API\n" + ex.Message);
            }

            return default;
        }

        /// <summary>
        /// Gets online members count in a distribution list.
        /// </summary>
        /// <param name="groupId">Distribution list group GUID.</param>
        /// <returns><see cref="Task{TResult}"/> Online members count in distribution list.</returns>
        [HttpGet]
        [Route("GetDistributionListMembersOnlineCount")]
        public async Task<int> GetDistributionListMembersOnlineCountAsync([FromQuery]string groupId)
        {
            try
            {
                int onlineUserCount;
                onlineUserCount = await this.presenceDataRepository.GetDistributionListMembersOnlineCountAsync(groupId);
                return onlineUserCount;
            }
            catch (MsalException ex)
            {
                this.telemetryClient.TrackTrace($"A Msal error occurred in GetDistributionListMembersOnlineCountAsync: {ex.Message}, Parameters: {groupId}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
                await this.HttpContext.Response.WriteAsync("An authentication error occurred while acquiring a token for downstream API\n" + ex.ErrorCode + "\n" + ex.Message);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in GetDistributionListMembersOnlineCountAsync: {ex.Message}, Parameters: {groupId}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.HttpContext.Response.ContentType = "text/plain";
                this.HttpContext.Response.StatusCode = (int)HttpStatusCode.InternalServerError;
                await this.HttpContext.Response.WriteAsync("An error occurred while calling the downstream API\n" + ex.Message);
            }

            return 0;
        }
    }
}