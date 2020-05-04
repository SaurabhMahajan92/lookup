// <copyright file="PeopleContactsData.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Models
{
    using Newtonsoft.Json;

    /// <summary>
    /// This model is for user contact data from Skype.
    /// </summary>
    public class PeopleContactsData
    {
        /// <summary>
        /// Gets or sets user principal name from AAD.
        /// </summary>
        [JsonProperty("uri")]
        public string UserPrincipalName { get; set; }
    }
}
