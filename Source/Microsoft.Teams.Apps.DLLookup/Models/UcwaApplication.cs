// <copyright file="UcwaApplication.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Models
{
    /// <summary>
    /// This model represents UCWA Application values required for getting presence.
    /// </summary>
    public class UcwaApplication
    {
        /// <summary>
        /// Gets or sets UCWA People Root Uri.
        /// </summary>
        public string UcwaPeopleRootUri { get; set; }

        /// <summary>
        /// Gets or sets result of UCWA application creation API call.
        /// </summary>
        public string CreateUcwaAppsResults { get; set; }

        /// <summary>
        /// Gets or sets access token which is generated with UCWA applications root Uri.
        /// </summary>
        public string AccessToken { get; set; }

        /// <summary>
        /// Gets or sets UCWA applications root Uri value.
        /// </summary>
        public string UcwaApplicationsRootUri { get; set; }
    }
}
