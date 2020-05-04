// <copyright file="MSGraphBatchResponse.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// This model represents the response of a batch request.
    /// </summary>
    /// <typeparam name="T">T type.</typeparam>
    public class MSGraphBatchResponse<T>
    {
        /// <summary>
        /// Gets or sets the Id of each request in batch call.
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets the body of each request in batch call.
        /// </summary>
        [JsonProperty("body")]
        public T Body { get; set; }
    }
}
