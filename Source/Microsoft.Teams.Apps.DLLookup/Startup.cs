// <copyright file="Startup.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup
{
    using Microsoft.AspNetCore.Builder;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.AspNetCore.SpaServices.ReactDevelopmentServer;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Identity.Web;
    using Microsoft.Identity.Web.Client.TokenCacheProviders;
    using Microsoft.Teams.Apps.DLLookup.Helpers.Extentions;

    /// <summary>
    /// Default Startup class.
    /// </summary>
    public class Startup
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Startup"/> class.
        /// </summary>
        /// <param name="configuration">Instance of application configuration properties.</param>
        public Startup(IConfiguration configuration)
        {
            this.Configuration = configuration;
        }

        /// <summary>
        /// Gets application configuration value.
        /// </summary>
        public IConfiguration Configuration { get; }

        /// <summary>
        /// This method gets called by the runtime. Use this method to add services to the container.
        /// </summary>
        /// <param name="services">IServiceCollection instance.</param>
        public void ConfigureServices(IServiceCollection services)
        {
            var scopes = this.Configuration["AzureAd:GraphScope"].Split(new char[] { ' ' }, System.StringSplitOptions.RemoveEmptyEntries);

            services.AddMemoryCache();

            services.AddApplicationInsightsTelemetry();

            services.AddProtectWebApiWithMicrosoftIdentityPlatformV2(this.Configuration)
                   .AddProtectedApiCallsWebApis(this.Configuration, scopes)
                   .AddInMemoryTokenCaches();

            services.AddMvc().SetCompatibilityVersion(CompatibilityVersion.Version_2_2);

            // In production, the React files will be served from this directory
            services.AddSpaStaticFiles(configuration =>
            {
                configuration.RootPath = "ClientApp/build";
            });

            services.AddRepositories();
            services.AddProtectedApiCallHelper();
        }

        /// <summary>
        /// This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        /// </summary>
        /// <param name="app">IApplicationBuilder instance.</param>
        /// <param name="env">IHostingEnvironment instance.</param>
        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        {
            app.UseSession();

            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseExceptionHandler("/Error");
                app.UseHsts();
            }

            app.UseHttpsRedirection();
            app.UseAuthentication();

            app.UseStaticFiles();
            app.UseSpaStaticFiles();

            app.UseMvc(routes =>
            {
                routes.MapRoute(
                    name: "default",
                    template: "{controller}/{action=Index}/{id?}");
            });

            app.UseSpa(spa =>
            {
                spa.Options.SourcePath = "ClientApp";

                if (env.IsDevelopment())
                {
                    spa.UseReactDevelopmentServer(npmScript: "start");
                }
            });
        }
    }
}
