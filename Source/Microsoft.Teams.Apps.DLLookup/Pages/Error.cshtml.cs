// <copyright file="Error.cshtml.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Pages
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.AspNetCore.Mvc.RazorPages;

    /// <summary>
    /// Creating ErrorModel class.
    /// </summary>
    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
#pragma warning disable SA1649 // File name should match first type name
    public class ErrorModel : PageModel
#pragma warning disable SA1649 // File name should match first type name
    {
        /// <summary>
        /// Gets or sets RequestId property.
        /// </summary>
        public string RequestId { get; set; }

        /// <summary>
        /// Gets a value indicating whether Request id is required to show or not.
        /// </summary>
        public bool ShowRequestId => !string.IsNullOrEmpty(this.RequestId);

        /// <summary>
        /// Set Request ID value.
        /// </summary>
        public void OnGet()
        {
            this.RequestId = Activity.Current?.Id ?? this.HttpContext.TraceIdentifier;
        }
    }
}
