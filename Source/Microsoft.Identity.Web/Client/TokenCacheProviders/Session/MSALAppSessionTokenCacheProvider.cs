// <copyright file="MSALAppSessionTokenCacheProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved. https://github.com/Azure-Samples/active-directory-dotnetcore-daemon-v2
// </copyright>

/************************************************************************************************
The MIT License (MIT)

Copyright (c) 2015 Microsoft Corporation

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
***********************************************************************************************/

namespace Microsoft.Identity.Web.Client.TokenCacheProviders
{
    using System;
    using System.Diagnostics;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authentication.AzureAD.UI;
    using Microsoft.AspNetCore.Http;
    using Microsoft.Extensions.Options;
    using Microsoft.Identity.Client;

    /// <summary>
    /// An implementation of token cache for Confidential clients backed by Http session.
    /// </summary>
    /// <seealso cref="https://aka.ms/msal-net-token-cache-serialization"/>
    public class MSALAppSessionTokenCacheProvider : IMSALAppTokenCacheProvider
    {
        /// <summary>
        /// The application cache key.
        /// </summary>
        internal string AppCacheId;

        /// <summary>
        /// Gets the HTTP context being used by this app.
        /// </summary>
        internal HttpContext HttpContext
        {
            get { return this.httpContextAccessor.HttpContext; }
        }

        /// <summary>
        /// HTTP context accessor.
        /// </summary>
        internal IHttpContextAccessor httpContextAccessor;

        /// <summary>
        /// The duration till the tokens are kept in memory cache. In production, a higher value , upto 90 days is recommended.
        /// </summary>
        private readonly DateTimeOffset cacheDuration = DateTimeOffset.Now.AddHours(12);

        private static ReaderWriterLockSlim sessionLock = new ReaderWriterLockSlim(LockRecursionPolicy.NoRecursion);

        /// <summary>
        /// The App's whose cache we are maintaining.
        /// </summary>
        private string appId;

        /// <summary>Initializes a new instance of the <see cref="MSALAppSessionTokenCacheProvider"/> class.</summary>
        /// <param name="azureAdOptionsAccessor">The azure ad options accessor.</param>
        /// <exception cref="ArgumentNullException">AzureADOptions - The app token cache needs {nameof(AzureADOptions)}.</exception>
        public MSALAppSessionTokenCacheProvider(IOptionsMonitor<AzureADOptions> azureAdOptionsAccessor, IHttpContextAccessor httpContextAccessor)
        {
            this.httpContextAccessor = httpContextAccessor;
            if (azureAdOptionsAccessor.CurrentValue == null && string.IsNullOrWhiteSpace(azureAdOptionsAccessor.CurrentValue.ClientId))
            {
                throw new ArgumentNullException(nameof(AzureADOptions), $"The app token cache needs {nameof(AzureADOptions)}, populated with clientId to initialize.");
            }

            this.appId = azureAdOptionsAccessor.CurrentValue.ClientId;
        }

        /// <summary>Initializes this instance of TokenCacheProvider with essentials to initialize themselves.</summary>
        /// <param name="tokenCache">The token cache instance of MSAL application.</param>
        /// <param name="httpcontext">The Httpcontext whose Session will be used for caching.This is required by some providers.</param>
        public void Initialize(ITokenCache tokenCache, HttpContext httpcontext)
        {
            this.AppCacheId = this.appId + "_AppTokenCache";

            tokenCache.SetBeforeAccessAsync(this.AppTokenCacheBeforeAccessNotificationAsync);
            tokenCache.SetAfterAccessAsync(this.AppTokenCacheAfterAccessNotificationAsync);
            tokenCache.SetBeforeWrite(this.AppTokenCacheBeforeWriteNotification);
        }

        /// <summary>
        /// if you want to ensure that no concurrent write take place, use this notification to place a lock on the entry.
        /// </summary>
        /// <param name="args">Contains parameters used by the MSAL call accessing the cache.</param>
        private void AppTokenCacheBeforeWriteNotification(TokenCacheNotificationArgs args)
        {
            // Since we are using a SessionCache ,whose methods are threads safe, we need not to do anything in this handler.
        }

        /// <summary>
        /// Clears the TokenCache's copy of this user's cache.
        /// </summary>
        public void Clear()
        {
            sessionLock.EnterWriteLock();
            try
            {
                Debug.WriteLine($"INFO: Clearing session {this.HttpContext.Session.Id}, cacheId {this.AppCacheId}");

                // Reflect changes in the persistent store
                this.HttpContext.Session.Remove(this.AppCacheId);
                this.HttpContext.Session.CommitAsync().Wait();
            }
            finally
            {
                sessionLock.ExitWriteLock();
            }
        }

        /// <summary>
        /// Triggered right before MSAL needs to access the cache. Reload the cache from the persistence store in case it changed since the last access.
        /// </summary>
        /// <param name="args">Contains parameters used by the MSAL call accessing the cache.</param>
        private async Task AppTokenCacheBeforeAccessNotificationAsync(TokenCacheNotificationArgs args)
        {
            await this.HttpContext.Session.LoadAsync();

            sessionLock.EnterReadLock();
            try
            {
                byte[] blob;
                if (this.HttpContext.Session.TryGetValue(this.AppCacheId, out blob))
                {
                    Debug.WriteLine($"INFO: Deserializing session {this.HttpContext.Session.Id}, cacheId {this.AppCacheId}");
                    args.TokenCache.DeserializeMsalV3(blob, shouldClearExistingCache: true);
                }
                else
                {
                    Debug.WriteLine($"INFO: cacheId {this.AppCacheId} not found in session {this.HttpContext.Session.Id}");
                }
            }
            finally
            {
                sessionLock.ExitReadLock();
            }
        }

        /// <summary>
        /// Triggered right after MSAL accessed the cache.
        /// </summary>
        /// <param name="args">Contains parameters used by the MSAL call accessing the cache.</param>
        private async Task AppTokenCacheAfterAccessNotificationAsync(TokenCacheNotificationArgs args)
        {
            // if the access operation resulted in a cache update
            if (args.HasStateChanged)
            {
                sessionLock.EnterWriteLock();
                try
                {
                    Debug.WriteLine($"INFO: Serializing session {this.HttpContext.Session.Id}, cacheId {this.AppCacheId}");

                    // Reflect changes in the persistent store
                    byte[] blob = args.TokenCache.SerializeMsalV3();
                    this.HttpContext.Session.Set(this.AppCacheId, blob);
                    await this.HttpContext.Session.CommitAsync();
                }
                finally
                {
                    sessionLock.ExitWriteLock();
                }
            }
        }
    }
}