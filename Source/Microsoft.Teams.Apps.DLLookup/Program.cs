// <copyright file="Program.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup
{
    using Microsoft.AspNetCore;
    using Microsoft.AspNetCore.Hosting;

    /// <summary>
    /// Default Program class.
    /// </summary>
    public class Program
    {
        /// <summary>
        /// Default Main method.
        /// </summary>
        /// <param name="args">string array input parameters.</param>
        public static void Main(string[] args)
        {
            CreateWebHostBuilder(args).Build().Run();
        }

        /// <summary>
        /// Method to create default builder.
        /// </summary>
        /// <param name="args">string input parameter from Main method.</param>
        /// <returns>Calls Startup method.</returns>
        public static IWebHostBuilder CreateWebHostBuilder(string[] args) =>
            WebHost.CreateDefaultBuilder(args)
                .UseStartup<Startup>();
    }
}
