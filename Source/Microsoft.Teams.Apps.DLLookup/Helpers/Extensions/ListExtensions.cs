// <copyright file="ListExtensions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Helpers.Extentions
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Class for List Extension methods.
    /// </summary>
    public static class ListExtensions
    {
        /// <summary>
        /// This method is to split Lists into smaller batches.
        /// </summary>
        /// <typeparam name="T">T type.</typeparam>
        /// <param name="source">Source List to split.</param>
        /// <param name="nSize">Size value to split the list with 40 as default value.</param>
        /// <returns>A <see cref="IEnumerable{TResult}"/> representing the sub-lists by specified size.</returns>
        public static IEnumerable<List<T>> SplitList<T>(this List<T> source, int nSize = 40)
        {
            for (int i = 0; i < source.Count; i += nSize)
            {
                yield return source.GetRange(i, Math.Min(nSize, source.Count - i));
            }
        }
    }
}
