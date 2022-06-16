// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

namespace Microsoft.PowerToys.Run.Plugin.Graph
{
    using System;

    public static class Util
    {
        public static bool ContainsIC(this string original, string value)
        {
            return original.Contains(value, StringComparison.OrdinalIgnoreCase);
        }
    }
}
