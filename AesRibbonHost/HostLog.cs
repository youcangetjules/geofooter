using System;
using System.IO;

namespace Aliniant.AesRibbonHost
{
    internal static class HostLog
    {
        private static readonly object Gate = new object();

        public static string LogPath
        {
            get
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "GeoFooter", "Logs");
                Directory.CreateDirectory(dir);
                return Path.Combine(dir, "AesRibbonHost.log");
            }
        }

        public static void Write(string message)
        {
            try
            {
                lock (Gate)
                {
                    File.AppendAllText(
                        LogPath,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " | " + message + Environment.NewLine);
                }
            }
            catch
            {
                // never throw from logging
            }
        }
    }
}
