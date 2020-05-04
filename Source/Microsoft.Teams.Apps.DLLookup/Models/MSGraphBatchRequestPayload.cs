// <copyright file="MSGraphBatchRequestPayload.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Models
{
    using System.Collections.Generic;

    /// <summary>
    /// This model class represents the batch request for MS Graph.
    /// </summary>
    public class MSGraphBatchRequestPayload
    {
        /// <summary>
        /// Gets or sets the list of requests in a batch call.
        /// </summary>
        public List<MSGraphBatchRequest> Requests { get; set; }
    }
}
