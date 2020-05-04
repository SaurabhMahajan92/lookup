// <copyright file="MSALPerUserMemoryTokenCacheProvider.cs" company="Microsoft">
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
    using System.Security.Claims;
    using Microsoft.AspNetCore.Http;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Identity.Client;

    /// <summary>
    /// An implementation of token cache for both Confidential and Public clients backed by MemoryCache.
    /// MemoryCache is useful in Api scenarios where there is no HttpContext.Session to cache data.
    /// </summary>
    /// <seealso cref="https://aka.ms/msal-net-token-cache-serialization"/>
    public class MSALPerUserMemoryTokenCacheProvider : IMSALUserTokenCacheProvider
    {
        /// <summary>
        /// The backing MemoryCache instance.
        /// </summary>
        internal IMemoryCache MemoryCache;

        private readonly MSALMemoryTokenCacheOptions cacheOptions;

        /// <summary>
        /// Enables the singleton object to access the right HttpContext.
        /// </summary>
        private IHttpContextAccessor httpContextAccessor;

        /// <summary>Initializes a new instance of the <see cref="MSALPerUserMemoryTokenCacheProvider"/> class.</summary>
        /// <param name="cache">The memory cache instance.</param>
        public MSALPerUserMemoryTokenCacheProvider(IMemoryCache cache, MSALMemoryTokenCacheOptions option, IHttpContextAccessor httpContextAccessor)
        {
            this.MemoryCache = cache;
            this.httpContextAccessor = httpContextAccessor;

            if (option == null)
            {
                this.cacheOptions = new MSALMemoryTokenCacheOptions();
            }
            else
            {
                this.cacheOptions = option;
            }
        }

        /// <summary>Initializes this instance of TokenCacheProvider with essentials to initialize themselves.</summary>
        /// <param name="tokenCache">The token cache instance of MSAL application.</param>
        /// <param name="httpcontext">The Httpcontext whose Session will be used for caching.This is required by some providers.</param>
        /// <param name="user">The signed-in user for whom the cache needs to be established. Not needed by all providers.</param>
        public void Initialize(ITokenCache tokenCache, HttpContext httpcontext, ClaimsPrincipal user)
        {
            tokenCache.SetBeforeAccess(this.UserTokenCacheBeforeAccessNotification);
            tokenCache.SetAfterAccess(this.UserTokenCacheAfterAccessNotification);
            tokenCache.SetBeforeWrite(this.UserTokenCacheBeforeWriteNotification);
        }

        /// <summary>
        /// Clears the TokenCache's copy of this user's cache.
        /// </summary>
        public void Clear(string accountId)
        {
            this.MemoryCache.Remove(accountId);
        }

        /// <summary>
        /// Triggered right after MSAL accessed the cache.
        /// </summary>
        /// <param name="args">Contains parameters used by the MSAL call accessing the cache.</param>
        private void UserTokenCacheAfterAccessNotification(TokenCacheNotificationArgs args)
        {
            // if the access operation resulted in a cache update
            if (args.HasStateChanged)
            {
                string cacheKey = this.httpContextAccessor.HttpContext.User.GetMsalAccountId();

                if (string.IsNullOrWhiteSpace(cacheKey))
                {
                    return;
                }

                // Ideally, methods that load and persist should be thread safe.MemoryCache.Get() is thread safe.
                this.MemoryCache.Set(cacheKey, args.TokenCache.SerializeMsalV3(), this.cacheOptions.SlidingExpiration);
            }
        }

        /// <summary>
        /// Triggered right before MSAL needs to access the cache. Reload the cache from the persistence store in case it
        /// changed since the last access.
        /// </summary>
        /// <param name="args">Contains parameters used by the MSAL call accessing the cache.</param>
        private void UserTokenCacheBeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            string cacheKey = this.httpContextAccessor.HttpContext.User.GetMsalAccountId();

            if (string.IsNullOrWhiteSpace(cacheKey))
            {
                return;
            }

            byte[] tokenCacheBytes = (byte[])this.MemoryCache.Get(cacheKey);
            args.TokenCache.DeserializeMsalV3(tokenCacheBytes, shouldClearExistingCache: true);
        }

        /// <summary>
        /// if you want to ensure that no concurrent write take place, use this notification to place a lock on the entry.
        /// </summary>
        /// <param name="args">Contains parameters used by the MSAL call accessing the cache.</param>
        private void UserTokenCacheBeforeWriteNotification(TokenCacheNotificationArgs args)
        {
        }
    }
}