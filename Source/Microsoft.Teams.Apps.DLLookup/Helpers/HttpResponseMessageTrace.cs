// <copyright file="HttpResponseMessageTrace.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
namespace Microsoft.Teams.Apps.DLLookup.Helpers
{
    using System.Net.Http;

    /// <summary>
    /// creating <see cref="HttpResponseMessageTrace"/> class.
    /// HttpResponseMessageTrace helps to streamline telemetry trace logs.
    /// </summary>
    public class HttpResponseMessageTrace
    {
        /// <summary>
        /// Gives Details of HttpResponseMessageTrace for logging purpose.
        /// </summary>
        /// <param name="response">HttpResponseMessage object.</param>
        /// <param name="customTraceInfo">Custom information to be logged.</param>
        /// <returns>A <see cref="string"/>representing the HTTP response message trace.</returns>
        public static string GetHttpResponseMessageTrace(HttpResponseMessage response, string customTraceInfo)
        {
            return $"HttpStatusCode : {response.StatusCode}. HttpMessage : {response.Content.ReadAsStringAsync().Result}. CustomTraceInfo : {customTraceInfo}";
        }
    }
}