using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Rainmeter;

namespace Info
{
    internal class Measure
    {
        private string Type;
        private string OnFinishAction;
        private string SkinName;
        private API api;

        internal Measure() { }

        internal void Reload(Rainmeter.API api, ref double maxValue)
        {
            this.api = api;
            Type = api.ReadString("Type", "").Trim();
            SkinName = api.ReadString("SkinName", "").Trim();
            OnFinishAction = api.ReadString("OnFinishAction", "").Trim();

            if (string.IsNullOrEmpty(Type))
            {
                api.Log(API.LogType.Error, "Info.dll: 'Type' must be specified in the measure.");
            }



        }
        internal double Update()
        {


            string skinsPath = api.ReplaceVariables("#SKINSPATH#").Trim();
            string skinFolder = Path.Combine(skinsPath, SkinName);

            if (!Directory.Exists(skinFolder))
            {
                api.Log(API.LogType.Error, $"Info.dll: Skin folder '{skinFolder}' does not exist.");
                return 0.0;
            }

            try
            {
                if (Type.Equals("Size", StringComparison.OrdinalIgnoreCase))
                {
                    long folderSize = GetDirectorySize(skinFolder);
                    double sizeInMb = folderSize / (1024.0 * 1024.0); // Convert to MB
                    return Math.Floor(sizeInMb); // Return value without decimals
                }
                else if (Type.Equals("DateModified", StringComparison.OrdinalIgnoreCase))
                {
                    DateTime lastModified = GetDirectoryLastModified(skinFolder);
                    return lastModified.ToOADate(); // Convert to a double for Rainmeter
                }
                else if (Type.Equals("FindClone", StringComparison.OrdinalIgnoreCase))
                {
                    string cloneFile = Path.Combine(skinFolder, "Clone.txt");
                    return File.Exists(cloneFile) ? 1.0 : 0.0; // Return 1 if file exists, 0 otherwise
                }

            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"Info.dll: Error processing Type '{Type}' for folder '{skinFolder}': {ex.Message}");
            }

            return 0.0; // Default return value for unsupported types or errors
        }

        internal string GetString()
        {
            if (string.IsNullOrEmpty(SkinName))
            {
                api.Log(API.LogType.Error, "Info.dll: 'SkinName' is not specified.");
                return string.Empty;
            }

            string skinsPath = api.ReplaceVariables("#SKINSPATH#").Trim();
            string skinFolder = Path.Combine(skinsPath, SkinName);

            if (!Directory.Exists(skinFolder))
            {
                api.Log(API.LogType.Error, $"Info.dll: Skin folder '{skinFolder}' does not exist.");
                return string.Empty;
            }

            try
            {
                if (Type.Equals("Size", StringComparison.OrdinalIgnoreCase))
                {
                    long folderSize = GetDirectorySize(skinFolder);
                    double sizeInMb = folderSize / (1024.0 * 1024.0); // Convert to MB
                    return $"{sizeInMb:F0} MB"; // Display without decimal
                }
                else if (Type.Equals("FindClone", StringComparison.OrdinalIgnoreCase))
                {
                    string cloneFile = Path.Combine(skinFolder, "Clone.txt");
                    return File.Exists(cloneFile) ? "1" : "0"; // Return "1" if found, otherwise "0"
                }
                else if (Type.Equals("DateModified", StringComparison.OrdinalIgnoreCase))
                {
                    DateTime lastModified = GetDirectoryLastModified(skinFolder);
                    return lastModified.ToString("yyyy-MM-dd HH:mm:ss");
                }
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"Info.dll: Error retrieving string for Type '{Type}': {ex.Message}");
            }

            return string.Empty; // Default return value for unsupported types
        }

        private long GetDirectorySize(string folderPath)
        {
            long size = 0;

            foreach (string file in Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories))
            {
                FileInfo fileInfo = new FileInfo(file);
                size += fileInfo.Length;
            }

            return size;
        }

        private DateTime GetDirectoryLastModified(string folderPath)
        {
            DateTime lastModified = Directory.GetLastWriteTime(folderPath);

            foreach (string file in Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories))
            {
                DateTime fileLastModified = File.GetLastWriteTime(file);
                if (fileLastModified > lastModified)
                {
                    lastModified = fileLastModified;
                }
            }

            return lastModified;
        }

        internal void Execute(string args)
        {
            if (string.IsNullOrEmpty(args))
            {
                api.Log(API.LogType.Error, "Info.dll: No arguments provided.");
                return;
            }

            string skinsPath = api.ReplaceVariables("#SKINSPATH#").Trim();
            string settingsPath = Path.Combine(api.ReplaceVariables("#SETTINGSPATH#").Trim(), "Rainmeter.ini");

            if (Type.Equals("Uninstall", StringComparison.OrdinalIgnoreCase))
            {
                HandleUninstall(skinsPath, settingsPath, args.Trim());
            }
            else if (Type.Equals("clearini", StringComparison.OrdinalIgnoreCase))
            {
                HandleClearINI(settingsPath, args.Trim());
            }
            else if (Type.Equals("Links", StringComparison.OrdinalIgnoreCase))
            {
                HandleLinks();
            }

            else if (Type.Equals("Clone", StringComparison.OrdinalIgnoreCase))
            {
                string[] splitArgs = args.Split('|');
                if (splitArgs.Length == 2)
                {
                    HandleClone(skinsPath, splitArgs[0].Trim(), splitArgs[1].Trim());
                }
                else
                {
                    api.Log(API.LogType.Error, "Info.dll: Invalid arguments for 'Clone'. Expected format: 'OriginalSkinName|NewSkinName'.");
                }
            }
            else
            {
                api.Log(API.LogType.Error, $"Info.dll: Unknown Type '{Type}' specified.");
            }

            if (!string.IsNullOrEmpty(OnFinishAction))
            {
                api.Execute(OnFinishAction);
            }
        }

        private void HandleLinks()
        {
            // Fetch DeviantArtLinks and Skin_Name variables
            string deviantArtLinks = api.ReplaceVariables("#DeviantArtLinks#").Trim();
            string skinName = api.ReplaceVariables("#Skin_Name#").Trim();

            if (string.IsNullOrEmpty(deviantArtLinks))
            {
                api.Log(API.LogType.Error, "Info.dll: 'DeviantArtLinks' is empty or not provided.");
            }

            if (string.IsNullOrEmpty(skinName))
            {
                api.Log(API.LogType.Error, "Info.dll: 'Skin_Name' is empty or not provided.");
            }

            if (string.IsNullOrEmpty(deviantArtLinks) || string.IsNullOrEmpty(skinName))
            {
                return;
            }

            // Split the DeviantArtLinks into link IDs and skin names
            string[] parts = deviantArtLinks.Split(':');
            if (parts.Length != 2)
            {
                api.Log(API.LogType.Error, "Info.dll: 'DeviantArtLinks' is not formatted correctly. Expected format: link1|link2|link3:Skin1|Skin2|Skin3");
                return;
            }

            string[] linkIds = parts[0].Split('|');
            string[] skinNames = parts[1].Split('|');

            if (linkIds.Length != skinNames.Length)
            {
                api.Log(API.LogType.Error, "Info.dll: Mismatch between the number of link IDs and skin names in 'DeviantArtLinks'.");
                return;
            }

            // Match the current skin name
            int index = Array.IndexOf(skinNames, skinName);
            if (index == -1)
            {
                api.Log(API.LogType.Warning, $"Info.dll: Skin name '{skinName}' not found in the skin names list.");
                return;
            }

            // Construct the DeviantArt URL
            string linkId = linkIds[index];
            string deviantArtUrl = $"https://www.deviantart.com/nstechbytes/art/{linkId}";

            api.Log(API.LogType.Debug, $"Info.dll: Constructed DeviantArt URL: {deviantArtUrl}");

            // Execute the URL in a browser or perform further actions
            api.Execute($"[\"{deviantArtUrl}\"]");
        }

        private void HandleUninstall(string skinsPath, string settingsPath, string skinName)
        {
            string skinFolder = Path.Combine(skinsPath, skinName);
            if (Directory.Exists(skinFolder))
            {
                try
                {
                    Directory.Delete(skinFolder, true);
                    api.Log(API.LogType.Debug, $"Info.dll: Skin folder '{skinFolder}' deleted.");
                }
                catch (Exception ex)
                {
                    api.Log(API.LogType.Error, $"Info.dll: Error deleting skin folder '{skinFolder}': {ex.Message}");
                }
            }
            else
            {
                api.Log(API.LogType.Warning, $"Info.dll: Skin folder '{skinFolder}' does not exist.");
            }

            if (File.Exists(settingsPath))
            {
                try
                {
                    var lines = File.ReadAllLines(settingsPath);
                    var updatedLines = ClearSkinSections(lines, skinName);

                    File.WriteAllLines(settingsPath, updatedLines);
                    api.Log(API.LogType.Debug, $"Info.dll: Removed all sections for '{skinName}' from settings file.");
                }
                catch (Exception ex)
                {
                    api.Log(API.LogType.Error, $"Info.dll: Error modifying settings file '{settingsPath}': {ex.Message}");
                }
            }
            else
            {
                api.Log(API.LogType.Warning, $"Info.dll: Settings file '{settingsPath}' does not exist.");
            }
        }

        private void HandleClearINI(string settingsPath, string skinName)
        {
            if (File.Exists(settingsPath))
            {
                try
                {
                    var lines = File.ReadAllLines(settingsPath);
                    var updatedLines = ClearSkinSections(lines, skinName);

                    File.WriteAllLines(settingsPath, updatedLines);
                    api.Log(API.LogType.Debug, $"Info.dll: Cleared all sections for '{skinName}' from settings file.");
                }
                catch (Exception ex)
                {
                    api.Log(API.LogType.Error, $"Info.dll: Error modifying settings file '{settingsPath}': {ex.Message}");
                }
            }
            else
            {
                api.Log(API.LogType.Warning, $"Info.dll: Settings file '{settingsPath}' does not exist.");
            }
        }


        private string[] ClearSkinSections(string[] lines, string skinName)
        {
            var updatedLines = new List<string>();
            bool inTargetSection = false;

            foreach (string line in lines)
            {
                // Detect the start of a section
                if (line.StartsWith("[") && line.EndsWith("]"))
                {
                    // Check if the section belongs to the target skin (both root and sub-sections)
                    inTargetSection =
                        line.Equals($"[{skinName}]", StringComparison.OrdinalIgnoreCase) ||  // Root section like [catppuccin]
                        line.StartsWith($"[{skinName}\\", StringComparison.OrdinalIgnoreCase); // Sub-sections like [catppuccin\...]
                }

                // Add lines to the updated list only if not in the target section
                if (!inTargetSection)
                {
                    updatedLines.Add(line);
                }
            }

            return updatedLines.ToArray();
        }



        private void HandleClone(string skinsPath, string originalSkinName, string newSkinName)
        {
            string originalSkinFolder = Path.Combine(skinsPath, originalSkinName);
            string newSkinFolder = Path.Combine(skinsPath, newSkinName);

            if (!Directory.Exists(originalSkinFolder))
            {
                api.Log(API.LogType.Error, $"Info.dll: Original skin folder '{originalSkinFolder}' does not exist.");
                return;
            }

            if (Directory.Exists(newSkinFolder))
            {
                api.Log(API.LogType.Warning, $"Info.dll: New skin folder '{newSkinFolder}' already exists.");
                return;
            }

            try
            {
                Directory.CreateDirectory(newSkinFolder);

                foreach (string dirPath in Directory.GetDirectories(originalSkinFolder, "*", SearchOption.AllDirectories))
                {
                    Directory.CreateDirectory(dirPath.Replace(originalSkinFolder, newSkinFolder));
                }

                foreach (string filePath in Directory.GetFiles(originalSkinFolder, "*.*", SearchOption.AllDirectories))
                {
                    File.Copy(filePath, filePath.Replace(originalSkinFolder, newSkinFolder), true);
                }

                // Create an empty Clone.txt file
                File.WriteAllText(Path.Combine(newSkinFolder, "Clone.txt"), string.Empty);
                api.Log(API.LogType.Debug, $"Info.dll: Skin '{originalSkinName}' cloned to '{newSkinName}' with Clone.txt created.");
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"Info.dll: Error cloning skin: {ex.Message}");
            }
        }
    }

    public static class Plugin
    {
        static IntPtr StringBuffer = IntPtr.Zero;

        [DllExport]
        public static void Initialize(ref IntPtr data, IntPtr rm)
        {
            data = GCHandle.ToIntPtr(GCHandle.Alloc(new Measure()));
        }

        [DllExport]
        public static void Finalize(IntPtr data)
        {
            GCHandle.FromIntPtr(data).Free();

            if (StringBuffer != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(StringBuffer);
                StringBuffer = IntPtr.Zero;
            }
        }

        [DllExport]
        public static void Reload(IntPtr data, IntPtr rm, ref double maxValue)
        {
            Measure measure = (Measure)GCHandle.FromIntPtr(data).Target;
            measure.Reload(new API(rm), ref maxValue);
        }

        [DllExport]
        public static double Update(IntPtr data)
        {
            Measure measure = (Measure)GCHandle.FromIntPtr(data).Target;
            return measure.Update();
        }

        [DllExport]
        public static IntPtr GetString(IntPtr data)
        {
            Measure measure = (Measure)GCHandle.FromIntPtr(data).Target;
            string value = measure.GetString();

            IntPtr stringBuffer = IntPtr.Zero;
            if (!string.IsNullOrEmpty(value))
            {
                stringBuffer = Marshal.StringToHGlobalUni(value);
            }

            return stringBuffer;
        }


        [DllExport]
        public static void ExecuteBang(IntPtr data, [MarshalAs(UnmanagedType.LPWStr)] string args)
        {
            Measure measure = (Measure)GCHandle.FromIntPtr(data).Target;
            measure.Execute(args);
        }
    }
}