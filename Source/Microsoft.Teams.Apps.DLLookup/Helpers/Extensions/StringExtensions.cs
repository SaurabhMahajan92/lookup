// <copyright file="StringExtensions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.DLLookup.Helpers.Extentions
{
    /// <summary>
    /// Class for String Extension methods.
    /// </summary>
    public static class StringExtensions
    {
        /// <summary>
        /// Truncates the string to the provided maximum length.
        /// </summary>
        /// <param name="value">String to truncate.</param>
        /// <param name="maxLength">How many chars to truncate.</param>
        /// <returns>Truncated string.</returns>
        public static string Truncate(this string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value))
            {
                return value;
            }

            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }
    }
}
