// <copyright file="ProtectedApiCallHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.AspNetCore.Http;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Identity.Client;
    using Microsoft.Identity.Web.Client;
    using Microsoft.Teams.Apps.DLLookup.Helpers;
    using Microsoft.Teams.Apps.DLLookup.Helpers.Extentions;
    using Microsoft.Teams.Apps.DLLookup.Models;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// <see cref="ProtectedApiCallHelper"/> class is used to call Web APIs with proper authentication.
    /// </summary>
    public class ProtectedApiCallHelper : IProtectedApiCallHelper
    {
        private static readonly HttpClient HttpClient = new HttpClient();
        private readonly ITokenAcquisition tokenAcquisition;
        private readonly IHttpContextAccessor httpContextAccessor;
        private readonly TelemetryClient telemetryClient;
        private readonly IConfiguration configuration;

        /// <summary>
        /// Initializes a new instance of the <see cref="ProtectedApiCallHelper"/> class.
        /// </summary>
        /// <param name="configuration">Singleton instance of application configuration.</param>
        /// <param name="tokenAcquisition">Scoped ITokenAcquisition instance for acquiring token.</param>
        /// <param name="httpContextAccessor">Scoped IHttpContextAccessor instance for giving access to HTTPContext.</param>
        /// <param name="telemetryClient">Singleton TelemetryClient instance used to send telemetry to Azure application insights.</param>
        public ProtectedApiCallHelper(IConfiguration configuration, ITokenAcquisition tokenAcquisition, IHttpContextAccessor httpContextAccessor, TelemetryClient telemetryClient)
        {
            this.tokenAcquisition = tokenAcquisition;
            this.httpContextAccessor = httpContextAccessor;
            this.telemetryClient = telemetryClient;
            this.configuration = configuration;
        }

        /// <summary>
        /// This method is to call batch web API on behalf of user.
        /// </summary>
        /// <param name="webApiUrl">Batch Web API URL to call.</param>
        /// <param name="batchWebApiUrls">Batch of Web API URLs to call.</param>
        /// <param name="accessToken">Access token for authentication.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public async Task<List<JObject>> CallBatchWebApiOnBehalfOfUserAsync(string webApiUrl, List<string> batchWebApiUrls, string accessToken)
        {
            List<JObject> peoplePresenceDataJObjectList = new List<JObject>();

            if (!string.IsNullOrEmpty(accessToken))
            {
                using MultipartContent batchContent = new MultipartContent("batching", Guid.NewGuid().ToString());

                foreach (string url in batchWebApiUrls)
                {
                    batchContent.Add(
                        new HttpMessageContent(
                            new HttpRequestMessage(
                            HttpMethod.Get,
                            url)));
                }

                using var request = new HttpRequestMessage(HttpMethod.Post, webApiUrl)
                {
                    Content = batchContent,
                };

                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("multipart/batching"));

                using HttpResponseMessage httpResponseMessageMultiPart = HttpClient.SendAsync(request).Result;

                if (httpResponseMessageMultiPart.IsSuccessStatusCode)
                {
                    MultipartMemoryStreamProvider multipartContent = await httpResponseMessageMultiPart.Content.ReadAsMultipartAsync();

                    foreach (HttpContent currentContent in multipartContent.Contents)
                    {
                        if (currentContent.Headers.ContentType.MediaType.Equals("application/http", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!currentContent.Headers.ContentType.Parameters.Any(parameter => parameter.Name.Equals("msgtype", StringComparison.OrdinalIgnoreCase) && parameter.Value.Equals("response", StringComparison.OrdinalIgnoreCase)))
                            {
                                currentContent.Headers.ContentType.Parameters.Add(new NameValueHeaderValue("msgtype", "response"));
                            }

                            using HttpResponseMessage httpResponseMessage = currentContent.ReadAsHttpResponseMessageAsync().Result;

                            if (httpResponseMessage.IsSuccessStatusCode)
                            {
                                peoplePresenceDataJObjectList.Add(JObject.Parse(httpResponseMessage.Content.ReadAsStringAsync().Result));
                            }
                            else if (httpResponseMessage.StatusCode == HttpStatusCode.TooManyRequests)
                            {
                                this.telemetryClient.TrackTrace($"HttpException:CallBatchWebApiOnBehalfOfUserAsync : Too many requests. Parameter: {string.Join(",", batchWebApiUrls)}", SeverityLevel.Error);
                                break;
                            }
                            else
                            {
                                this.telemetryClient.TrackTrace(
                                        HttpResponseMessageTrace.GetHttpResponseMessageTrace(
                                            httpResponseMessage,
                                            $"HttpException:CallBatchWebApiOnBehalfOfUserAsync.WebApiUrl: {string.Join(",", batchWebApiUrls).Truncate(9500)}"),
                                        SeverityLevel.Error);
                            }
                        }
                    }
                }
                else if (httpResponseMessageMultiPart.StatusCode == HttpStatusCode.TooManyRequests)
                {
                    this.telemetryClient.TrackTrace($"BatchHttpException:CallBatchWebApiOnBehalfOfUserAsync : Too many requests. Parameter: {string.Join(",", batchWebApiUrls)}", SeverityLevel.Error);
                }
                else
                {
                    this.telemetryClient.TrackTrace(
                        HttpResponseMessageTrace.GetHttpResponseMessageTrace(
                            httpResponseMessageMultiPart,
                            $"BatchHttpException:CallBatchWebApiOnBehalfOfUserAsync.WebApiUrl: {string.Join(",", batchWebApiUrls)}"),
                        SeverityLevel.Error);
                }
            }

            return peoplePresenceDataJObjectList;
        }

        /// <summary>
        /// This method is to call Graph API on behalf of user.
        /// </summary>
        /// <typeparam name="T">T type.</typeparam>
        /// <param name="scopes">Graph scopes.</param>
        /// <param name="query">Graph query.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public async Task<List<T>> CallGraphApiOnBehalfOfUser<T>(string[] scopes, string query)
        {
            // We use MSAL.NET to get a token to call the API On Behalf Of the current user
            try
            {
                List<T> members = new List<T>();
                string accessToken = await this.tokenAcquisition.GetAccessTokenOnBehalfOfUser(this.httpContextAccessor.HttpContext, scopes, this.configuration["AzureAd:TenantId"]);
                string resultString = await this.CallGraphApiOnBehalfOfUser(accessToken, query);

                if (this.GetValue<List<T>>(resultString, "value") != null)
                {
                    members.AddRange(this.GetValue<List<T>>(resultString, "value"));
                    query = this.GetValue<string>(resultString, "@odata.nextLink");

                    while (query != null)
                    {
                        resultString = await this.CallGraphApiOnBehalfOfUser(accessToken, query);
                        query = this.GetValue<string>(resultString, "@odata.nextLink");
                        if (this.GetValue<List<T>>(resultString, "value") != null)
                        {
                            members.AddRange(this.GetValue<List<T>>(resultString, "value"));
                        }
                    }
                }

                return members;
            }
            catch (MsalUiRequiredException ex)
            {
                this.telemetryClient.TrackTrace($"A Msal error occurred in CallGraphApiOnBehalfOfUser: {ex.Message}" + " Parameter:" + query, SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.tokenAcquisition.ReplyForbiddenWithWwwAuthenticateHeader(this.httpContextAccessor.HttpContext, scopes, ex);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in CallGraphApiOnBehalfOfUser:  {ex.Message}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
            }

            return default;
        }

        /// <summary>
        /// This method is to call Graph API on behalf of user with Direct Token.
        /// </summary>
        /// <typeparam name="T">T type.</typeparam>
        /// <param name="scopes">Graph scopes.</param>
        /// <param name="query">Graph query.</param>
        /// <param name="selectDirectJToken">Boolean to select JToken from json response or response itself.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public async Task<T> CallGraphApiOnBehalfOfUserWithDirectJToken<T>(string[] scopes, string query, bool selectDirectJToken = false)
        {
            // We use MSAL.NET to get a token to call the API On Behalf Of the current user
            try
            {
                string accessToken = await this.tokenAcquisition.GetAccessTokenOnBehalfOfUser(this.httpContextAccessor.HttpContext, scopes, this.configuration["AzureAd:TenantId"]);
                string resultString = await this.CallGraphApiOnBehalfOfUser(accessToken, query);
                List<T> members = new List<T>();

                if (selectDirectJToken)
                {
                    return JsonConvert.DeserializeObject<T>(resultString);
                }
                else
                {
                    return this.GetValue<T>(resultString, "value");
                }
            }
            catch (MsalUiRequiredException ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in CallGraphApiOnBehalfOfUserList: {ex.Message}" + "Parameter:" + query, SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.tokenAcquisition.ReplyForbiddenWithWwwAuthenticateHeader(this.httpContextAccessor.HttpContext, scopes, ex);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in CallGraphApiOnBehalfOfUserWithDirectJToken:  {ex.Message}", SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
            }

            return default;
        }

        /// <summary>
        /// This method is to call Web API on behalf of User.
        /// </summary>
        /// <param name="webApiUrl">Web API URL to call.</param>
        /// <param name="accessToken">Access token for authentication.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public async Task<JObject> CallWebApiOnBehalfOfUserAsync(string webApiUrl, string accessToken)
        {
            if (!string.IsNullOrEmpty(accessToken))
            {
                // Setup request
                using var request = new HttpRequestMessage(HttpMethod.Get, webApiUrl);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);

                using HttpResponseMessage response = await HttpClient.SendAsync(request);
                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    JObject result = JObject.Parse(json);
                    return result;
                }
                else
                {
                    this.telemetryClient.TrackTrace(
                        HttpResponseMessageTrace.GetHttpResponseMessageTrace(
                            response,
                            $"CallWebApiOnBehalfOfUserAsync.WebApiUrl: {webApiUrl}"),
                        SeverityLevel.Error);
                }
            }

            return default;
        }

        /// <summary>
        /// To call GraphApi on behalf of user using payload.
        /// </summary>
        /// <param name="scopes">Graph scopes.</param>
        /// <param name="requestUrl">Graph API URL.</param>
        /// <param name="payload">Input JSON.</param>
        /// <returns>A <see cref="T"/> representing the result of the asynchronous operation.</returns>
        public async Task<dynamic> CallGraphApiPostOnBehalfOfUser(string[] scopes, string requestUrl, string payload)
        {
            string accessToken = await this.tokenAcquisition.GetAccessTokenOnBehalfOfUser(this.httpContextAccessor.HttpContext, scopes, this.configuration["AzureAd:TenantId"]);

            using var request = new HttpRequestMessage(HttpMethod.Post, requestUrl)
            {
                Content = new StringContent(payload, Encoding.UTF8, "application/json"),
            };
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            using HttpResponseMessage response = await HttpClient.SendAsync(request);
            string content = await response.Content.ReadAsStringAsync();
            if (response.IsSuccessStatusCode)
            {
                return this.GetValue<List<MSGraphBatchResponse<dynamic>>>(content, "responses");
            }
            else
            {
                this.telemetryClient.TrackTrace(
                    HttpResponseMessageTrace.GetHttpResponseMessageTrace(
                        response,
                        $"CallGraphApiPostOnBehalfOfUser.requestUrl:{requestUrl}. CallGraphApiPostOnBehalfOfUser.payload:{payload}."),
                    SeverityLevel.Error);
            }

            return content;
        }
       
        /// <summary>
        /// This method is to call Web API on behalf of user with payload.
        /// </summary>
        /// <param name="webApiUrl">Web API URL to call.</param>
        /// <param name="accessToken">Access Token for authentication.</param>
        /// <param name="payload">Dynamic input as payload.</param>
        /// <returns>A <see cref="Task{TResult}"/> representing the result of the asynchronous operation.</returns>
        public async Task<string> CallWebApiOnBehalfOfUserWithJsonPayloadAsync(string webApiUrl, string accessToken, dynamic payload)
        {
            if (!string.IsNullOrEmpty(accessToken))
            {
                using var request = new HttpRequestMessage(HttpMethod.Post, webApiUrl)
                {
                    Content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json"),
                };

                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                using HttpResponseMessage httpResponseMessage = HttpClient.SendAsync(request).Result;

                if (httpResponseMessage.IsSuccessStatusCode)
                {
                    return await httpResponseMessage.Content.ReadAsStringAsync();
                }
                else
                {
                    this.telemetryClient.TrackTrace(
                        HttpResponseMessageTrace.GetHttpResponseMessageTrace(
                            httpResponseMessage,
                            $"CallWebApiOnBehalfOfUserWithJsonPayloadAsync.WebApiUrl: {webApiUrl}. CallWebApiOnBehalfOfUserWithJsonPayLoadAsync.payLoad: {payload}"),
                        SeverityLevel.Error);
                }
            }

            return default;
        }

        /// <summary>
        /// This method is to get access token used for authentication.
        /// </summary>
        /// <param name="scopes">Scopes or resources.</param>
        /// <returns>A <see cref="Task{string}"/> representing the access token.</returns>
        public async Task<string> GetAccessToken(string[] scopes)
        {
            // We use MSAL.NET to get a token to call the API On Behalf Of the current user
            try
            {
                string accessToken = await this.tokenAcquisition.GetAccessTokenOnBehalfOfUser(this.httpContextAccessor.HttpContext, scopes, this.configuration["AzureAd:TenantId"]);
                return accessToken;
            }
            catch (MsalUiRequiredException ex)
            {
                this.telemetryClient.TrackTrace($"An error occurred in CallGraphApiOnBehalfOfUser: {ex.Message}" + " Parameter:" + scopes, SeverityLevel.Error);
                this.telemetryClient.TrackException(ex);
                this.tokenAcquisition.ReplyForbiddenWithWwwAuthenticateHeader(this.httpContextAccessor.HttpContext, scopes, ex);
            }

            return default;
        }

        /// <summary>
        /// To get value from JSON.
        /// </summary>
        /// <typeparam name="T">T type.</typeparam>
        /// <param name="json">Input JSON value.</param>
        /// <param name="jsonPropertyName">Property to retrieve.</param>
        /// <returns>A <see cref="{TResult}"/> representing the value from JSON.</returns>
        public T GetValue<T>(string json, string jsonPropertyName)
        {
            if (!string.IsNullOrEmpty(json))
            {
                JObject parsedResult = JObject.Parse(json);
                if (parsedResult[jsonPropertyName] != null)
                {
                    return parsedResult[jsonPropertyName].ToObject<T>();
                }
                else
                {
                    return default;
                }
            }
            else
            {
                return default;
            }
        }

        /// <summary>
        /// To get value from JSON.
        /// </summary>
        /// <typeparam name="T">T type.</typeparam>
        /// <param name="json">Dynamic input parameter.</param>
        /// <param name="jsonPropertyName">Json Property Name.</param>
        /// <returns>A <see cref="{TResult}"/> representing the value from JSON.</returns>
        public T GetValue<T>(dynamic json, string jsonPropertyName)
        {
            if (!(json is null))
            {
                return json.SelectToken(jsonPropertyName).ToObject<T>();
            }
            else
            {
                return default;
            }
        }

        /// <summary>
        /// Executes Graph API request.
        /// </summary>
        /// <param name="accessToken">Access token to authenticate Graph API call.</param>
        /// <param name="query">Graph API query.</param>
        /// <returns>Response JSON.</returns>
        private async Task<string> CallGraphApiOnBehalfOfUser(string accessToken, string query)
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, query);

            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            using HttpResponseMessage response = await HttpClient.SendAsync(request);
            string content = await response.Content.ReadAsStringAsync();
            if (response.IsSuccessStatusCode)
            {
                return content;
            }
            else
            {
                this.telemetryClient.TrackTrace(
                    HttpResponseMessageTrace.GetHttpResponseMessageTrace(
                        response,
                        $"CallGraphApiOnBehalfOfUser.query:{query}"),
                    SeverityLevel.Error);
            }

            return content;
        }
    }
}
