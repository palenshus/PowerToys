// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

namespace Microsoft.PowerToys.Run.Plugin.OneNote
{
    public class OneNoteResult
    {
        public OneNoteResult(string title)
        {
            this.Title = title;
        }

        public string Title { get; }

        public override string ToString()
        {
            return this.Title;
        }
    }
}
