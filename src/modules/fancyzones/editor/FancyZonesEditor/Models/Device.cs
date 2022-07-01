// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Globalization;
using System.Text;
using System.Windows;

namespace FancyZonesEditor.Utils
{
    public class Device
    {
        public string MonitorName { get; set; }

        public Size MonitorSize { get; set; }

        public string VirtualDesktopId { get; set; }

        public Rect ScaledBounds { get; private set; }

        public Rect WorkAreaRect { get; private set; }

        public int Dpi { get; set; }

        public Device(string monitorName, string virtualDesktopId, int dpi, Rect workArea, Size monitorSize)
        {
            MonitorName = monitorName;
            VirtualDesktopId = virtualDesktopId;
            Dpi = dpi;
            WorkAreaRect = workArea;
            MonitorSize = monitorSize;
        }

        public Device(Rect workArea, Size monitorSize)
        {
            WorkAreaRect = workArea;
            MonitorSize = monitorSize;
        }

        public void Scale(double scaleFactor)
        {
            WorkAreaRect = new Rect(Math.Round(WorkAreaRect.X * scaleFactor), Math.Round(WorkAreaRect.Y * scaleFactor), Math.Round(WorkAreaRect.Width * scaleFactor), Math.Round(WorkAreaRect.Height * scaleFactor));
        }

        public override string ToString()
        {
            var sb = new StringBuilder();

            sb.AppendFormat(CultureInfo.InvariantCulture, "MonitorName: {0}{1}", MonitorName, Environment.NewLine);
            sb.AppendFormat(CultureInfo.InvariantCulture, "Virtual desktop: {0}{1}", VirtualDesktopId, Environment.NewLine);
            sb.AppendFormat(CultureInfo.InvariantCulture, "DPI: {0}{1}", Dpi, Environment.NewLine);

            string monitorSize = MonitorSize.ToString(CultureInfo.InvariantCulture);
            string workArea = string.Format(CultureInfo.InvariantCulture, "({0}, {1}, {2}, {3})", WorkAreaRect.X, WorkAreaRect.Y, WorkAreaRect.Width, WorkAreaRect.Height);

            sb.AppendFormat(CultureInfo.InvariantCulture, "Monitor size: {0}{1}", monitorSize, Environment.NewLine);
            sb.AppendFormat(CultureInfo.InvariantCulture, "Work area: {0}{1}", workArea, Environment.NewLine);

            return sb.ToString();
        }
    }
}
