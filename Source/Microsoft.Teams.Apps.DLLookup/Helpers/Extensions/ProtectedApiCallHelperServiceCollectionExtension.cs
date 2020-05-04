// <copyright file="ProtectedApiCallHelperServiceCollectionExtension.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Helpers.Extentions
{
    using Microsoft.Extensions.DependencyInjection;

    /// <summary>
    /// Class to add scoped service for ProtectedApiCallHelper class.
    /// </summary>
    public static class ProtectedApiCallHelperServiceCollectionExtension
    {
        /// <summary>
        /// Extension method to register ProtectedApiCallHelper service in DI container.
        /// </summary>
        /// <param name="services">IServiceCollection instance to which ProtectedApiCallHelper scoped service to be added in.</param>
        public static void AddProtectedApiCallHelper(this IServiceCollection services)
        {
            services.AddScoped<IProtectedApiCallHelper, ProtectedApiCallHelper>();
        }
    }
}
