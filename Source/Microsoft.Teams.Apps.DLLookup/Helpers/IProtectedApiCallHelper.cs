// <copyright file="IProtectedApiCallHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// interface <see cref="IProtectedApiCallHelper"/>.
    /// </summary>
    public interface IProtectedApiCallHelper
    {
        /// <summary>
        /// This method is to call batch web API on behalf of user.
        /// </summary>
        /// <param name="webApiUrl">Batch Web API URL to call.</param>
        /// <param name="batchWebApiUrls">Batch of Web API URLs to call.</param>
        /// <param name="accessToken">Access token for authentication.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        Task<List<JObject>> CallBatchWebApiOnBehalfOfUserAsync(string webApiUrl, List<string> batchWebApiUrls, string accessToken);

        /// <summary>
        /// This method is to call Graph API on behalf of user.
        /// </summary>
        /// <typeparam name="T">T type.</typeparam>
        /// <param name="scopes">Graph scopes.</param>
        /// <param name="query">Graph query.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        Task<List<T>> CallGraphApiOnBehalfOfUser<T>(string[] scopes, string query);

        /// <summary>
        /// This method is to call Graph API on behalf of user with Direct Token.
        /// </summary>
        /// <typeparam name="T">T type.</typeparam>
        /// <param name="scopes">Graph scopes.</param>
        /// <param name="query">Graph query.</param>
        /// <param name="selectDirectJToken">Boolean to select JToken from json response or response itself.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        Task<T> CallGraphApiOnBehalfOfUserWithDirectJToken<T>(string[] scopes, string query, bool selectDirectJToken = false);

        /// <summary>
        /// This method is to call Web API on behalf of User.
        /// </summary>
        /// <param name="webApiUrl">Web API URL to call.</param>
        /// <param name="accessToken">Access token for authentication.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        Task<JObject> CallWebApiOnBehalfOfUserAsync(string webApiUrl, string accessToken);

        /// <summary>
        /// To call GraphApi on behalf of user using payload.
        /// </summary>
        /// <param name="scopes">Graph scopes.</param>
        /// <param name="requestUrl">Graph API URL.</param>
        /// <param name="payload">Input JSON.</param>
        /// <returns>A <see cref="T"/> representing the result of the asynchronous operation.</returns>
        Task<dynamic> CallGraphApiPostOnBehalfOfUser(string[] scopes, string requestUrl, string payload);

        /// <summary>
        /// This method is to call Web API on behalf of user with payload.
        /// </summary>
        /// <param name="webApiUrl">Web API URL to call.</param>
        /// <param name="accessToken">Access Token for authentication.</param>
        /// <param name="payload">Dynamic input as payload.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        Task<string> CallWebApiOnBehalfOfUserWithJsonPayloadAsync(string webApiUrl, string accessToken, dynamic payLoad);

        /// <summary>
        /// This method is to get access token used for authentication.
        /// </summary>
        /// <param name="scopes">Scopes or resources.</param>
        /// <returns>A <see cref="Task{string}"/> representing the access token.</returns>
        Task<string> GetAccessToken(string[] scopes);

        /// <summary>
        /// To get value from JSON.
        /// </summary>
        /// <typeparam name="T">T type.</typeparam>
        /// <param name="json">Input JSON value.</param>
        /// <param name="jsonPropertyName">Property to retrieve.</param>
        /// <returns>A <see cref="{TResult}"/> representing the value from JSON.</returns>
        T GetValue<T>(dynamic json, string jsonPropertyName);

        /// <summary>
        /// To get value from JSON.
        /// </summary>
        /// <typeparam name="T">T type.</typeparam>
        /// <param name="json">Dynamic input parameter.</param>
        /// <param name="jsonPropertyName">Json Property Name.</param>
        /// <returns>A <see cref="{TResult}"/> representing the value from JSON.</returns>
        T GetValue<T>(string json, string jsonPropertyName);
    }
}