// <copyright file="PresenceAppAgentObject.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Models
{
    /// <summary>
    /// This model represent user agent information used for creating UCWA application.
    /// </summary>
    public class PresenceAppAgentObject
    {
        /// <summary>
        /// Gets or sets user agent for UCWA application.
        /// </summary>
        public string UserAgent { get; set; }

        /// <summary>
        /// Gets or sets endpoint Id for UCWA application.
        /// </summary>
        public string EndpointId { get; set; }

        /// <summary>
        /// Gets or sets Culture for UCWA application.
        /// </summary>
        public string Culture { get; set; }
    }
}
