// <copyright file="PeoplePresenceData.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Models
{
    using System;
    using System.Collections.Generic;
    using Newtonsoft.Json;

    /// <summary>
    /// This model is for member's presence information.
    /// </summary>
    public class PeoplePresenceData
    {
        /// <summary>
        /// Gets or sets the member's availability.
        /// </summary>
        public string Availability { get; set; }

        /// <summary>
        /// Gets value of sort order based on availability.
        /// </summary>
        public int AvailabilitySortOrder
        {
            get
            {
                int sortOrder = 6;
                switch (!string.IsNullOrEmpty(this.Availability) ? this.Availability.ToLower() : string.Empty)
                {
                    case "online":
                        sortOrder = 0;
                        break;
                    case "busy":
                        sortOrder = 1;
                        break;
                    case "donotdisturb":
                        sortOrder = 2;
                        break;
                    case "berightback":
                        sortOrder = 3;
                        break;
                    case "away":
                        sortOrder = 4;
                        break;
                    case "offline":
                        sortOrder = 5;
                        break;
                }

                return sortOrder;
            }
        }

        /// <summary>
        /// Gets or sets member's user principal name.
        /// </summary>
        public string UserPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets User AAD Id.
        /// </summary>
        public string Id { get; set; }
    }
}
