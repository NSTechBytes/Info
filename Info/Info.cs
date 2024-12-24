/*
       
//=====================================================================================================================================//
//                              Below Lines Help to use Plugin with Measures                                                           //
//=====================================================================================================================================//
[Devian_Links]
Measure=Plugin
Plugin=Info
Type=Links

[Size_Helper]
Measure=Plugin 
Plugin=Info 
Type=Size
UpdateDivider=-1

[Date_Helper]
Measure=Plugin 
Plugin=Info 
Type=DateModified
UpdateDivider=-1

[ClearIni_Helper]
Measure=Plugin 
Plugin=Info 
Type=ClearIni
[Gen_Structure]
Measure=Plugin 
Plugin=Info 
Type=GenerateStructure

[PatchNote]
Measure=Plugin 
Plugin=Info 
Type=PatchNote
NoMatchAction=

[Clone_Helper]
Measure=Plugin 
Plugin=Info 
Type=Clone 
OnFinishAction=

For Execution ([!CommandMeasure Clone_Helper "Skin that want to clone| new clone name"])

[Uinstall_Helper]
Measure=Plugin 
Plugin=Info 
Type=Uninstall 
OnFinishAction=

For Execution ([!CommandMeasure Uinstall_Helper "Skiv Name"])
  */
//=====================================================================================================================================//
//                                             Main  Code Start Here                                                                   //
//=====================================================================================================================================//
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Rainmeter;

namespace Info
{
    internal class Measure
    {
        private string Type;
        private string OnFinishAction;
        private string SkinName;
        private string SkinVer;
        private API api;
        private string NoMatchAction;
        
        internal Measure() { }

        internal void Reload(Rainmeter.API api, ref double maxValue)
        {
            this.api = api;
            Type = api.ReadString("Type", "").Trim();
            SkinName = api.ReplaceVariables("#Skin_Name#");
            SkinVer = api.ReplaceVariables("#Version#");
            OnFinishAction = api.ReadString("OnFinishAction", "").Trim();
            NoMatchAction = api.ReadString("NoMatchAction", "").Trim();

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
                    double sizeInMb = folderSize / (1024.0 * 1024.0); 
                    return Math.Floor(sizeInMb); 
                }
                else if (Type.Equals("DateModified", StringComparison.OrdinalIgnoreCase))
                {
                    DateTime lastModified = GetDirectoryLastModified(skinFolder);
                    return lastModified.ToOADate(); 
                }
                else if (Type.Equals("FindClone", StringComparison.OrdinalIgnoreCase))
                {
                    string cloneFile = Path.Combine(skinFolder, "Clone.txt");
                    return File.Exists(cloneFile) ? 1.0 : 0.0;
                }
                else if (Type.Equals("PatchNote", StringComparison.OrdinalIgnoreCase))
                {
                    ValidateUserNameAndCopyFonts();
                }

                else if (Type.Equals("GenerateStructure", StringComparison.OrdinalIgnoreCase))
                {
                    CheckAndCopyStructure();


                }
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"Info.dll: Error processing Type '{Type}' for folder '{skinFolder}': {ex.Message}");
            }

            return 0.0;
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
                    double sizeInMb = folderSize / (1024.0 * 1024.0);
                    return $"{sizeInMb:F0} MB"; 
                }
                else if (Type.Equals("FindClone", StringComparison.OrdinalIgnoreCase))
                {
                    string cloneFile = Path.Combine(skinFolder, "Clone.txt");
                    return File.Exists(cloneFile) ? "1" : "0"; 
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

            return string.Empty; 
        }
        //=====================================================================================================================================//
        //                                            GetDirectory Size                                                                        //
        //=====================================================================================================================================//
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
        //=====================================================================================================================================//
        //                                            GenerateStructure                                                                        //
        //=====================================================================================================================================//

        private void CheckAndCopyStructure()
        {
            try
            {
              
                string nekDataPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Rainmeter", "NekData");
                string skinResourcePath = Path.Combine(api.ReplaceVariables("#SKINSPATH#"), SkinName, "@Resources", "@Structure");
                string nekDataSkinPath = Path.Combine(nekDataPath, SkinName);
                string versionFilePath = Path.Combine(nekDataSkinPath, $"{SkinVer}.txt");

              
                if (!Directory.Exists(nekDataPath))
                {
                    Directory.CreateDirectory(nekDataPath);
                    api.Log(API.LogType.Debug, "Created NekData folder at: " + nekDataPath);
                }

               
                if (!Directory.Exists(skinResourcePath))
                {
                    api.Log(API.LogType.Error, $"Structure folder not found in the @Resources folder of '{SkinName}'.");
                    return;
                }

               
                if (!Directory.Exists(nekDataSkinPath))
                {
                    Directory.CreateDirectory(nekDataSkinPath);
                    api.Log(API.LogType.Debug, "Created skin folder in NekData: " + nekDataSkinPath);
                }

                
                string skinVersion = api.ReadString("Version", "").Trim();
                bool copyStructure = true;

                if (File.Exists(versionFilePath))
                {
                    api.Log(API.LogType.Debug, $"Version file found: {versionFilePath}. No need to copy structure.");
                    copyStructure = false;
                }
                else
                {
                    api.Log(API.LogType.Debug, $"Version file not found: {versionFilePath}. Copying structure...");
                }

               
                if (copyStructure)
                {
                    CopyDirectory(skinResourcePath, nekDataSkinPath);

                  
                    File.WriteAllText(versionFilePath, skinVersion);
                    api.Log(API.LogType.Debug, "Copied Structure folder and created version file: " + versionFilePath);
                }
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, "Error in CheckAndCopyStructure: " + ex.Message);
            }
        }
        //=====================================================================================================================================//
        //                                            Copy Helper                                                                              //
        //=====================================================================================================================================//
        private void CopyDirectory(string sourceDir, string destDir)
        {
            Directory.CreateDirectory(destDir);
            foreach (var file in Directory.GetFiles(sourceDir))
            {
                string destFile = Path.Combine(destDir, Path.GetFileName(file));
                File.Copy(file, destFile, overwrite: true);
            }
            foreach (var dir in Directory.GetDirectories(sourceDir))
            {
                string destSubDir = Path.Combine(destDir, Path.GetFileName(dir));
                CopyDirectory(dir, destSubDir);
            }
        }

        //=====================================================================================================================================//
        //                                            PatchNote                                                                                //
        //=====================================================================================================================================//

        private void ValidateUserNameAndCopyFonts()
        {
            try
            {
                string currentUsername = Environment.UserName.Trim();
                string configuredUsername = api.ReplaceVariables("#UserName#");

                if (!currentUsername.Equals(configuredUsername, StringComparison.OrdinalIgnoreCase))
                {
                   
                    string skinResourcePath = Path.Combine(api.ReplaceVariables("#SKINSPATH#"), SkinName, "@Resources");
                    string patchNoteFilePath = Path.Combine(skinResourcePath, "PatchNote.nek");
                    string fontsSourcePath = Path.Combine(skinResourcePath, "Fonts");
                    string nekStartFontsPath = Path.Combine(api.ReplaceVariables("#SKINSPATH#"), "#NekStart", "@Resources", "Fonts");

                  
                    string patchNoteContent = "[Variables]\nUserName = " + currentUsername;
                    File.WriteAllText(patchNoteFilePath, patchNoteContent);
                    api.Log(API.LogType.Debug, $"Wrote username '{currentUsername}' to PatchNote.nek.");

                  
                    if (Directory.Exists(fontsSourcePath))
                    {
                        CopyDirectory(nekStartFontsPath, fontsSourcePath);
                        api.Execute(NoMatchAction);
                        api.Log(API.LogType.Debug, "Copied Fonts folder to NekStart @Resources.");
                    }
                    else
                    {
                        api.Log(API.LogType.Error, "Fonts folder not found in the skin's @Resources directory.");
                    }
                }
                else
                {
                    api.Log(API.LogType.Debug, "Username matches the system's username. No action required.");
                }
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, "Error in ValidateUserNameAndCopyFonts: " + ex.Message);
            }
        }
        //=====================================================================================================================================//
        //                                            DateModified                                                                             //
        //=====================================================================================================================================//

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
        //=====================================================================================================================================//
        //                                             DeviantArtLinks                                                                         //
        //=====================================================================================================================================//
       public void HandleLinks()
        {
           
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

            
            int index = Array.IndexOf(skinNames, skinName);
            if (index == -1)
            {
                api.Log(API.LogType.Warning, $"Info.dll: Skin name '{skinName}' not found in the skin names list.");
                return;
            }

        
            string linkId = linkIds[index];
            string deviantArtUrl = $"https://www.deviantart.com/nstechbytes/art/{linkId}";

            api.Log(API.LogType.Debug, $"Info.dll: Constructed DeviantArt URL: {deviantArtUrl}");

          
            api.Execute($"[\"{deviantArtUrl}\"]");
        }
        //=====================================================================================================================================//
        //                                             UnInstall                                                                              //
        //=====================================================================================================================================//

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
        //=====================================================================================================================================//
        //                                            ClearIni                                                                                 //
        //=====================================================================================================================================//
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
               
                if (line.StartsWith("[") && line.EndsWith("]"))
                {
                   
                    inTargetSection =
                        line.Equals($"[{skinName}]", StringComparison.OrdinalIgnoreCase) || 
                        line.StartsWith($"[{skinName}\\", StringComparison.OrdinalIgnoreCase); 
                }

               
                if (!inTargetSection)
                {
                    updatedLines.Add(line);
                }
            }

            return updatedLines.ToArray();
        }


        //=====================================================================================================================================//
        //                                            Clone                                                                                    //
        //=====================================================================================================================================//
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

               
                File.WriteAllText(Path.Combine(newSkinFolder, "Clone.txt"), string.Empty);
                api.Log(API.LogType.Debug, $"Info.dll: Skin '{originalSkinName}' cloned to '{newSkinName}' with Clone.txt created.");
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"Info.dll: Error cloning skin: {ex.Message}");
            }
        }
    }
    //=====================================================================================================================================//
    //                                            Rainmeter Class                                                                          //
    //=====================================================================================================================================//
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