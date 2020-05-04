// <copyright file="MSGraphBatchRequest.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Models
{
    /// <summary>
    /// This model represents individual request in a batch to be sent MS Graph.
    /// </summary>
    public class MSGraphBatchRequest
    {
        /// <summary>
        /// Gets or sets the unique id of each request in batch call.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets the HTTP method for each request in batch call.
        /// </summary>
        public string Method { get; set; }

        /// <summary>
        /// Gets or sets the URL of each request in batch call.
        /// </summary>
        public string URL { get; set; }
    }
}
