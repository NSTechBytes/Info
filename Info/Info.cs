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

For Execution ([!CommandMeasure Uinstall_Helper "Skin Name"])
  */
//=====================================================================================================================================//
//                                             Main  Code Start Here                                                                   //
//=====================================================================================================================================//
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic.Devices;
using Microsoft.Win32;
using Rainmeter;
using static Info.Measure.TaskbarUtils;

namespace Info
{
    internal class Measure
    {
        private string Type;
        private string OnFinishAction;
        private string SkinName;
        private string SkinVer;
        private Rainmeter.API api;
        private string NoMatchAction;
        private PerformanceCounter cpuCounter;
        private float lastCpuUsage = -1.0f;
        private DateTime lastUpdated = DateTime.MinValue;

        internal Measure() { }

        internal void Reload(Rainmeter.API api, ref double maxValue)
        {
            this.api = api;
            Type = api.ReadString("Type", "").Trim();
            SkinName = api.ReplaceVariables("#Skin_Name#");
            SkinVer = api.ReplaceVariables("#Version#");
            OnFinishAction = api.ReadString("OnFinishAction", "").Trim();
            NoMatchAction = api.ReadString("NoMatchAction", "").Trim();

          

            //    string[] validTypes = new string[] { "Size", "DateModified", "FindClone", "PatchNote", "GenerateStructure", "RamGB", "ChromeInstalled", "WebViewInstalled" };


            if (string.IsNullOrEmpty(Type))
            {
                Error("Info.dll: 'Type' must be specified in the measure.");
            }
            // else if (Array.IndexOf(validTypes, Type) == -1)
            //    {
            //     
            //         Error($"Info.dll: Invalid Type '{Type}' specified in the measure.");
            //     }
        }


        internal double Update()
        {
            string skinsPath = api.ReplaceVariables("#SKINSPATH#").Trim();
            string skinFolder = Path.Combine(skinsPath, SkinName);
            cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");

            if (!Directory.Exists(skinFolder))
            {
                Error($"Info.dll: Skin folder '{skinFolder}' does not exist.");
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
                else if (Type.Equals("RamGB", StringComparison.OrdinalIgnoreCase))
                {
                    return GetTotalRamInGB();
                }
                else if (Type.Equals("RamMB", StringComparison.OrdinalIgnoreCase))
                {
                    return GetTotalRamInMB();
                }
                else if (Type.Equals("PhysicalCores", StringComparison.OrdinalIgnoreCase))
                {
                    return GetTotalPhysicalCores();
                }
                else if (Type.Equals("OSBits", StringComparison.OrdinalIgnoreCase))
                {
                    return GetOSBits().Equals("64-bit") ? 64.0 : 32.0;
                }
                else if (Type.Equals("WindowsVersion", StringComparison.OrdinalIgnoreCase))
                {
                    string version = GetWindowsVersion();
                    Debug($"Detected Windows Version: {version}");
                    return version.Equals("Windows 7") ? 7.0
                        : version.Equals("Windows 8") ? 8.0
                        : version.Equals("Windows 8.1") ? 8.1
                        : version.Equals("Windows 10") ? 10.0
                        : version.Equals("Windows 11") ? 11.0
                        : 0.0;
                }
                else if (Type.Equals("ScreenWidth", StringComparison.OrdinalIgnoreCase))
                {
                    int screenWidth = GetScreenWidth();
                    Debug($"Detected Screen Width: {screenWidth}px");
                    return screenWidth;
                }
                else if (Type.Equals("ScreenHeight", StringComparison.OrdinalIgnoreCase))
                {
                    int screenHeight = GetScreenHeight();
                    Debug($"Detected Screen Height: {screenHeight}px");
                    return screenHeight;
                }
                else if (Type.Equals("MousePositionX", StringComparison.OrdinalIgnoreCase))
                {
                    int mouseX = GetMousePositionX();
                    Debug($"Detected Mouse X Position: {mouseX}px");
                    return mouseX;
                }
                else if (Type.Equals("MousePositionY", StringComparison.OrdinalIgnoreCase))
                {
                    int mouseY = GetMousePositionY();
                    Debug($"Detected Mouse Y Position: {mouseY}px");
                    return mouseY;
                }
                else if (Type.Equals("TaskbarPosition", StringComparison.OrdinalIgnoreCase))
                {
                    TaskbarPosition taskbarPosition = TaskbarUtils.GetTaskbarPosition();
                    Debug($"Detected Taskbar Position: {taskbarPosition}");
                    return (int)taskbarPosition;
                }
                else if (Type.Equals("InternetConnection", StringComparison.OrdinalIgnoreCase))
                {
                    bool isConnected = IsInternetAvailable();
                    Debug(
                        $"Internet Connection Status: {(isConnected ? "Connected" : "Disconnected")}"
                    );
                    return isConnected ? 1.0 : 0.0;
                }
                else if (Type.Equals("ChromeInstalled", StringComparison.OrdinalIgnoreCase))
                {
                    bool isChromeInstalled = IsChromeInstalled();
                    Debug(
                        $"Chrome Installation Status: {(isChromeInstalled ? "Installed" : "Not Installed")}"
                    );
                    return isChromeInstalled ? 1.0 : 0.0;
                }
                else if (Type.Equals("DayNightStatus", StringComparison.OrdinalIgnoreCase))
                {
                    double dayOrNight = GetDayOrNight();
                    api.Log(API.LogType.Debug, $"Day or Night status: {(dayOrNight == 1.0 ? "Day" : "Night")}");
                    return dayOrNight;
                }
                else if (Type.Equals("IdleTime", StringComparison.OrdinalIgnoreCase))
                {
                    int idleTimeInSeconds = GetIdleTimeInSeconds();
                    api.Log(API.LogType.Debug, $"Idle Time in Seconds: {idleTimeInSeconds}");
                    return idleTimeInSeconds;
                }
                else if (Type.Equals("PageSize", StringComparison.OrdinalIgnoreCase))
                {
                    uint pageSize = GetPageSize();
                    api.Log(API.LogType.Debug, $"Page Size: {pageSize} bytes");
                    return pageSize; // Optionally convert to MB or GB if required
                }
                else if (Type.Equals("CpuUsage", StringComparison.OrdinalIgnoreCase))
                {
                    float cpuUsage = GetCpuUsage();
                    api.Log(API.LogType.Debug, $"CPU Usage: {cpuUsage}%");
                    return cpuUsage; // Return the CPU usage percentage
                }
                else if (Type.Equals("RamUsage", StringComparison.OrdinalIgnoreCase))
                {
                    float ramUsage = GetRamUsage();
                    api.Log(API.LogType.Debug, $"RAM Usage: {ramUsage}%");
                    return ramUsage; // Return RAM usage percentage
                }




                else if (Type.Equals("WebViewInstalled", StringComparison.OrdinalIgnoreCase))
                {
                    bool isWebViewInstalled = IsWebViewInstalled();
                    Debug(
                        $"WebView2 Installation Status: {(isWebViewInstalled ? "Installed" : "Not Installed")}"
                    );
                    return isWebViewInstalled ? 1.0 : 0.0;
                }
            }
            catch (Exception ex)
            {
                Error(
                    $"Info.dll: Error processing Type '{Type}' for folder '{skinFolder}': {ex.Message}"
                );
            }

            return 0.0;
        }

        //=====================================================================================================================================//
        //                                            Logs and Execute Helper                                                                  //
        //=====================================================================================================================================//
        private void Error(string message)
        {
            api.Log(API.LogType.Error, $"Info.dll: {message}");
        }

        private void Debug(string message)
        {
            api.Log(API.LogType.Debug, $"Info.dll: {message}");
        }

        private void warninig(string message)
        {
            Debug($"Info.dll: {message}");
        }

        private void rmbang(string action)
        {
            api.Execute(action);
        }

        private void Log(string message)
        {
            api.Log(API.LogType.Notice, $"Info.dll: {message}");
        }

        //=====================================================================================================================================//
        //                                           Execute Function                                                                          //
        //=====================================================================================================================================//
        internal void Execute(string args)
        {
            if (string.IsNullOrEmpty(args))
            {
                Error("Info.dll: No arguments provided.");
                return;
            }

            string skinsPath = api.ReplaceVariables("#SKINSPATH#").Trim();
            string settingsPath = Path.Combine(
                api.ReplaceVariables("#SETTINGSPATH#").Trim(),
                "Rainmeter.ini"
            );

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
                    Error(
                        "Info.dll: Invalid arguments for 'Clone'. Expected format: 'OriginalSkinName|NewSkinName'."
                    );
                }
            }
            else
            {
                Error($"Info.dll: Unknown Type '{Type}' specified.");
            }

            if (!string.IsNullOrEmpty(OnFinishAction))
            {
                api.Execute(OnFinishAction);
            }
        }

        //=====================================================================================================================================//
        //                                           Get String Helper                                                                         //
        //=====================================================================================================================================//
        internal string GetString()
        {
            if (string.IsNullOrEmpty(SkinName))
            {
                Error("Info.dll: 'SkinName' is not specified.");
                return string.Empty;
            }

            string skinsPath = api.ReplaceVariables("#SKINSPATH#").Trim();
            string skinFolder = Path.Combine(skinsPath, SkinName);

            if (!Directory.Exists(skinFolder))
            {
                Error($"Info.dll: Skin folder '{skinFolder}' does not exist.");
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
                else if (Type.Equals("FullWindowsVersion", StringComparison.OrdinalIgnoreCase))
                {
                    string fullVersion = GetWindowsFullVersion();
                    Debug($"Detected Full Windows Version: {fullVersion}");
                    return fullVersion;
                }
                else if (Type.Equals("Username", StringComparison.OrdinalIgnoreCase))
                {
                    string username = GetUsername();
                    Debug($"Detected Username: {username}");
                    return username;
                }
                else if (Type.Equals("ComputerName", StringComparison.OrdinalIgnoreCase))
                {
                    string computerName = GetComputerName();
                    Debug($"Detected Computer Name: {computerName}");
                    return computerName;
                }
                else if (Type.Equals("IPAddress", StringComparison.OrdinalIgnoreCase))
                {
                    string ipAddress = GetLocalIPAddress();
                    Debug($"Detected IP Address: {ipAddress}");
                    return ipAddress;
                }
                else if (Type.Equals("TaskbarPosition_String", StringComparison.OrdinalIgnoreCase))
                {
                    string taskbarPosition = TaskbarUtils.GetTaskbarPositionString();
                    Debug($"Detected Taskbar Position: {taskbarPosition}");
                    return taskbarPosition;
                }
                else if (Type.Equals("TimeZoneStandard", StringComparison.OrdinalIgnoreCase))
                {
                    string timeZone = GetTimeZoneStandard();
                    api.Log(API.LogType.Debug, $"Time Zone Standard: {timeZone}");
                    return timeZone; // Optionally return a value to indicate successful execution
                }
                else if (Type.Equals("HostName", StringComparison.OrdinalIgnoreCase))
                {
                    string hostName = GetHostName();
                    api.Log(API.LogType.Debug, $"Host Name: {hostName}");
                    return hostName; // Optionally return a value to indicate successful execution
                }
                else if (Type.Equals("DomainName", StringComparison.OrdinalIgnoreCase))
                {
                    string domainName = GetDomainName();
                    api.Log(API.LogType.Debug, $"Domain Name: {domainName}");
                    return domainName; // Optionally return a value to indicate successful execution
                }
                else if (Type.Equals("UserSID", StringComparison.OrdinalIgnoreCase))
                {
                    string userSID = GetUserSID();
                    api.Log(API.LogType.Debug, $"User SID: {userSID}");
                    return userSID; // Optionally return a value to indicate successful execution
                }
                else if (Type.Equals("CurrentLanguage", StringComparison.OrdinalIgnoreCase))
                {
                    string currentLanguage = GetCurrentLanguage();
                    api.Log(API.LogType.Debug, $"Current Language: {currentLanguage}");
                    return currentLanguage; // Optionally return a value to indicate successful execution
                }
                else if (Type.Equals("CursorColor", StringComparison.OrdinalIgnoreCase))
                {
                    Color color = GetColorUnderMouseCursor();
                    if (!color.IsEmpty)
                    {
                        string rgbColor = $"{color.R},{color.G},{color.B}";
                        api.Log(API.LogType.Debug, $"Color under mouse cursor: {rgbColor}");
                        return rgbColor; // Returns the color in ARGB format as an integer
                    }
                    return "Unknown"; // Return 0 if there's an error
                }
                else if (Type.Equals("Uptime", StringComparison.OrdinalIgnoreCase))
                {
                    string uptime = GetSystemUptime();
                    api.Log(API.LogType.Debug, $"System Uptime: {uptime}");
                    return uptime; // Returning 0.0 as uptime is a string, but logged in debug
                }
                else if (Type.Equals("ActiveWindowTitle", StringComparison.OrdinalIgnoreCase))
                {
                    string activeWindowTitle = WindowHelper.GetActiveWindowTitle();
                    api.Log(API.LogType.Debug, $"Active Window Title: {activeWindowTitle}");
                    return activeWindowTitle; // Return the active window title
                }
                else if (Type.Equals("LastSleepTime", StringComparison.OrdinalIgnoreCase))
                {
                    DateTime lastSleepTime = SleepHelper.GetLastSleepTime();

                    if (lastSleepTime == DateTime.MinValue)
                    {
                        api.Log(API.LogType.Debug, "No sleep event found.");
                        return "No sleep event found";
                    }

                    api.Log(API.LogType.Debug, $"Last sleep time: {lastSleepTime}");
                    return lastSleepTime.ToString("yyyy-MM-dd HH:mm:ss"); // Return last sleep time in readable format
                }
                else if (Type.Equals("LastWakeTime", StringComparison.OrdinalIgnoreCase))
                {
                    DateTime lastWakeTime = WakeHelper.GetLastWakeTime();

                    if (lastWakeTime == DateTime.MinValue)
                    {
                        api.Log(API.LogType.Debug, "No wake event found.");
                        return "No wake event found";
                    }

                    api.Log(API.LogType.Debug, $"Last wake time: {lastWakeTime}");
                    return lastWakeTime.ToString("yyyy-MM-dd HH:mm:ss"); // Return last wake time in readable format
                }
                else if (Type.Equals("PressedKey", StringComparison.OrdinalIgnoreCase))
                {
                    string pressedKey = KeyboardHelper.GetPressedKey();
                    api.Log(API.LogType.Debug, $"Pressed key: {pressedKey}");
                    return pressedKey; // Return the pressed key as a string
                }







            }
            catch (Exception ex)
            {
                Error($"Info.dll: Error retrieving string for Type '{Type}': {ex.Message}");
            }

            return string.Empty;
        }
        //=====================================================================================================================================//
        //                                PressedKey                                                                                           //
        //====================================================================================================================================//


public class KeyboardHelper
    {
        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int vKey);

        public static string GetPressedKey()
        {
            for (int i = 0; i < 256; i++)
            {
                short keyState = GetAsyncKeyState(i);
                if ((keyState & 0x8000) != 0)
                {
                    return ((Keys)i).ToString(); // Return the name of the key that is pressed
                }
            }
            return "No key pressed";
        }
    }




    //=====================================================================================================================================//
    //                                 LastWakeTime                                                                                        //
    //====================================================================================================================================//

    public class WakeHelper
    {
        private static DateTime _lastWakeTime = DateTime.MinValue; // Cache for the last wake time
        private static DateTime _lastCheckedTime = DateTime.MinValue;
        private static readonly TimeSpan CacheDuration = TimeSpan.FromSeconds(60); // Cache duration, adjust as needed

        // Function to get the last wake time from the event logs (optimized)
        public static DateTime GetLastWakeTime()
        {
            try
            {
                // Update the cached value periodically (every 5 minutes, for example)
                if (DateTime.Now - _lastCheckedTime > CacheDuration)
                {
                    UpdateLastWakeTime();
                }

                // Return the cached last wake time
                return _lastWakeTime;
            }
            catch (Exception ex)
            {
                // Log error if something goes wrong
                //api.Log(API.LogType.Error, $"Error retrieving last wake time: {ex.Message}");
                return DateTime.MinValue; // Return a default value if an error occurs
            }
        }

        // Function to update the last wake time by querying the event logs
        private static void UpdateLastWakeTime()
        {
            try
            {
                // Get the System Event Log
                EventLog systemLog = new EventLog("System");

                // Filter for wake events (Event ID 1)
                var wakeEvents = systemLog.Entries.Cast<EventLogEntry>()
                    .Where(e => e.InstanceId == 1) // Event ID 1 indicates wake
                    .OrderByDescending(e => e.TimeGenerated)
                    .ToList();

                // If no wake event is found, return
                if (wakeEvents.Count == 0)
                {
                    _lastWakeTime = DateTime.MinValue;
                    return;
                }

                // Get the last wake time (most recent event with ID 1)
                _lastWakeTime = wakeEvents.First().TimeGenerated;
                _lastCheckedTime = DateTime.Now; // Update the last checked time
            }
            catch (Exception ex)
            {
                // Log error if something goes wrong
               // api.Log(API.LogType.Error, $"Error updating last wake time: {ex.Message}");
            }
        }
    }




    //=====================================================================================================================================//
    //                                 LastSleepTime                                                                                       //
    //=====================================================================================================================================//





    public class SleepHelper
    {
        private static DateTime _lastSleepTime = DateTime.MinValue; // Cache for the last sleep time
        private static DateTime _lastCheckedTime = DateTime.MinValue;
        private static readonly TimeSpan CacheDuration = TimeSpan.FromSeconds(60); // Cache duration, adjust as needed

        // Function to get the last sleep time from the event logs (optimized)
        public static DateTime GetLastSleepTime()
        {
            try
            {
                // Update the cached value periodically (every 5 minutes, for example)
                if (DateTime.Now - _lastCheckedTime > CacheDuration)
                {
                    UpdateLastSleepTime();
                }

                // Return the cached last sleep time
                return _lastSleepTime;
            }
            catch (Exception ex)
            {
                // Log error if something goes wrong
             //   api.Log(API.LogType.Error, $"Error retrieving last sleep time: {ex.Message}");
                return DateTime.MinValue; // Return a default value if an error occurs
            }
        }

        // Function to update the last sleep time by querying the event logs
        private static void UpdateLastSleepTime()
        {
            try
            {
                // Get the System Event Log
                EventLog systemLog = new EventLog("System");

                // Filter for sleep events (Event ID 42)
                var sleepEvents = systemLog.Entries.Cast<EventLogEntry>()
                    .Where(e => e.InstanceId == 42) // Event ID 42 indicates sleep
                    .OrderByDescending(e => e.TimeGenerated)
                    .ToList();

                // If no sleep event is found, return
                if (sleepEvents.Count == 0)
                {
                    _lastSleepTime = DateTime.MinValue;
                    return;
                }

                // Get the last sleep time (most recent event with ID 42)
                _lastSleepTime = sleepEvents.First().TimeGenerated;
                _lastCheckedTime = DateTime.Now; // Update the last checked time
            }
            catch (Exception ex)
            {
                // Log error if something goes wrong
                //api.Log(API.LogType.Error, $"Error updating last sleep time: {ex.Message}");
            }
        }
}




        //=====================================================================================================================================//
        //                                  UpTime                                                                                             //
        //=====================================================================================================================================//





        private string GetSystemUptime()
    {
        try
        {
            // Get the uptime in milliseconds from the Windows API (ulong)
            ulong uptimeMilliseconds = GetTickCount64();

            // Cast ulong to long since TimeSpan.FromMilliseconds expects a long
            long uptimeMillisecondsLong = (long)uptimeMilliseconds;

            // Convert milliseconds to a TimeSpan
            TimeSpan uptime = TimeSpan.FromMilliseconds(uptimeMillisecondsLong);

            // Format the TimeSpan into a readable string
            string formattedUptime = string.Format(
                "{0} days, {1} hours, {2} minutes, {3} seconds",
                uptime.Days,
                uptime.Hours,
                uptime.Minutes,
                uptime.Seconds
            );

            return formattedUptime;
        }
        catch (Exception ex)
        {
            api.Log(API.LogType.Error, $"Error retrieving system uptime: {ex.Message}");
            return "Error";
        }
    }

    [DllImport("kernel32.dll")]
    private static extern ulong GetTickCount64();






        //=====================================================================================================================================//
        //                                ActiveWindowTitle                                                                                    //
        //=====================================================================================================================================//




public class WindowHelper
    {
        // Importing necessary Windows APIs
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        // Function to get the title of the currently active window
        public static string GetActiveWindowTitle()
        {
            try
            {
                IntPtr hwnd = GetForegroundWindow(); // Get the handle of the currently active window

                if (hwnd == IntPtr.Zero)
                {
                    return "No active window"; // If no active window
                }

                StringBuilder windowTitle = new StringBuilder(256);
                int length = GetWindowText(hwnd, windowTitle, windowTitle.Capacity);

                // If no title is retrieved, return a default message
                if (length == 0)
                {
                    return "No title available";
                }

                // Return the window title (just the title, no full path)
                return windowTitle.ToString();
            }
            catch (Exception ex)
            {
                // Log errors if any occur
              //  api.Log(API.LogType.Error, $"Error retrieving active window title: {ex.Message}");
                return "Error retrieving title";
            }
        }
}




        //=====================================================================================================================================//
        //                                  Cursor Color                                                                                       //
        //=====================================================================================================================================//

        private Color GetColorUnderMouseCursor()
        {
            try
            {
                // Get the current mouse cursor position
                POINT cursorPosition;
                if (GetCursorPos(out cursorPosition))
                {
                    // Get the device context of the screen
                    IntPtr hdc = GetDC(IntPtr.Zero);

                    // Retrieve the color of the pixel under the mouse cursor
                    uint pixelColor = GetPixel(hdc, cursorPosition.X, cursorPosition.Y);

                    // Release the device context
                    ReleaseDC(IntPtr.Zero, hdc);

                    // Convert the color to a .NET Color object
                    int red = (int)(pixelColor & 0x000000FF);
                    int green = (int)((pixelColor & 0x0000FF00) >> 8);
                    int blue = (int)((pixelColor & 0x00FF0000) >> 16);

                    return Color.FromArgb(red, green, blue);
                }
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"Error retrieving color under mouse cursor: {ex.Message}");
            }

            return Color.Empty; // Return an empty color if there's an error
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("gdi32.dll")]
        private static extern uint GetPixel(IntPtr hdc, int nXPos, int nYPos);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);



        //=====================================================================================================================================//
        //                                  Ram Usage                                                                                          //
        //=====================================================================================================================================//


        private float GetRamUsage()
        {
            try
            {
                // Create a PerformanceCounter instance to measure available memory
                using (PerformanceCounter memoryCounter = new PerformanceCounter("Memory", "Available MBytes"))
                {
                    // Total physical memory
                    float totalMemory = GetTotalPhysicalMemoryInMB();

                    // Available memory in MB
                    float availableMemory = memoryCounter.NextValue();

                    // Calculate used memory
                    float usedMemory = totalMemory - availableMemory;

                    // Return RAM usage as a percentage
                    return (usedMemory / totalMemory) * 100.0f; // Percentage of used RAM
                }
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"Error retrieving RAM usage: {ex.Message}");
                return -1.0f; // Indicate an error
            }
        }

        private float GetTotalPhysicalMemoryInMB()
        {
            try
            {
                // Use the Microsoft.VisualBasic.Devices.ComputerInfo to retrieve total physical memory
                Microsoft.VisualBasic.Devices.ComputerInfo computerInfo = new Microsoft.VisualBasic.Devices.ComputerInfo();
                return computerInfo.TotalPhysicalMemory / (1024.0f * 1024.0f); // Convert bytes to MB
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"Error retrieving total physical memory: {ex.Message}");
                return 0.0f; // Indicate an error
            }
        }


        //=====================================================================================================================================//
        //                                  CPU Usage                                                                                          //
        //=====================================================================================================================================//



 

    private float GetCpuUsage()
    {
        try
        {
            // Update the CPU usage only once every second
            if (DateTime.Now.Subtract(lastUpdated).TotalSeconds >= 1)
            {
                // Wait for a second before getting the next value
                cpuCounter.NextValue(); // First call, no value yet
                Thread.Sleep(100); // Small sleep to get an accurate result

                // Now get the second value
                lastCpuUsage = cpuCounter.NextValue();
                lastUpdated = DateTime.Now;
            }

            return lastCpuUsage;
        }
        catch (Exception ex)
        {
            api.Log(API.LogType.Error, $"Error retrieving CPU usage: {ex.Message}");
            return -1.0f; // Indicate an error
        }
    }



    //=====================================================================================================================================//
    //                                   Page Size                                                                                         //
    //=====================================================================================================================================//


    private uint GetPageSize()
        {
            try
            {
                SYSTEM_INFO systemInfo = new SYSTEM_INFO();
                GetSystemInfo(ref systemInfo);

                // Return the page size in bytes
                return systemInfo.dwPageSize;
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"Error retrieving page size: {ex.Message}");
                return 0; // Indicate an error
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SYSTEM_INFO
        {
            public ushort wProcessorArchitecture;
            public ushort wReserved;
            public uint dwPageSize;
            public IntPtr lpMinimumApplicationAddress;
            public IntPtr lpMaximumApplicationAddress;
            public IntPtr dwActiveProcessorMask;
            public uint dwNumberOfProcessors;
            public uint dwProcessorType;
            public uint dwAllocationGranularity;
            public ushort wProcessorLevel;
            public ushort wProcessorRevision;
        }

        [DllImport("kernel32.dll")]
        private static extern void GetSystemInfo(ref SYSTEM_INFO lpSystemInfo);


        //=====================================================================================================================================//
        //                                    Idle Time                                                                                        //
        //=====================================================================================================================================//


        private int GetIdleTimeInSeconds()
        {
            try
            {
                LASTINPUTINFO lastInputInfo = new LASTINPUTINFO();
                lastInputInfo.cbSize = Marshal.SizeOf<LASTINPUTINFO>();

                if (GetLastInputInfo(ref lastInputInfo))
                {
                    uint idleTime = (uint)Environment.TickCount - lastInputInfo.dwTime;
                    return (int)(idleTime / 1000); // Convert milliseconds to seconds
                }
                else
                {
                    api.Log(API.LogType.Error, "Error calling GetLastInputInfo.");
                    return -1; // Indicate an error
                }
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"Error retrieving idle time: {ex.Message}");
                return -1; // Indicate an error
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct LASTINPUTINFO
        {
            public int cbSize;
            public uint dwTime;
        }

        [DllImport("user32.dll")]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);


        //=====================================================================================================================================//
        //                                    Current Language                                                                                 //
        //=====================================================================================================================================//



        private string GetCurrentLanguage()
        {
            try
            {
                // Get the current language of the computer
                string language = CultureInfo.CurrentCulture.DisplayName; // e.g., "English (United States)"
                return language;
            }
            catch (Exception ex)
            {
                // Log error if something goes wrong
                api.Log(API.LogType.Error, $"Error retrieving current language: {ex.Message}");
                return "Unknown";
            }
        }


        //=====================================================================================================================================//
        //                                    Host Name                                                                                        //
        //=====================================================================================================================================//
        private string GetHostName()
        {
            try
            {
                // Get the machine name (hostname)
                return Environment.MachineName;
            }
            catch (Exception ex)
            {
                // Log error if something goes wrong
                api.Log(API.LogType.Error, $"Error retrieving hostname: {ex.Message}");
                return "Unknown";
            }
        }
        //=====================================================================================================================================//
        //                                    User SID                                                                                         //
        //=====================================================================================================================================//
        private string GetUserSID()
        {
            try
            {
                // Get the current user's SID
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                return identity.User.Value; // Returns the SID as a string
            }
            catch (Exception ex)
            {
                // Log error if something goes wrong
                api.Log(API.LogType.Error, $"Error retrieving user SID: {ex.Message}");
                return "Unknown";
            }
        }

        //=====================================================================================================================================//
        //                                    Domain Name                                                                                      //
        //=====================================================================================================================================//
        private string GetDomainName()
        {
            try
            {
                // Retrieve the domain name of the machine
                string domainName = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;

                // Return the domain name (if available) or "None" if not part of a domain
                return string.IsNullOrEmpty(domainName) ? "None" : domainName;
            }
            catch (Exception ex)
            {
                // Log error if something goes wrong
                api.Log(API.LogType.Error, $"Error retrieving domain name: {ex.Message}");
                return "Unknown";
            }
        }
        //=====================================================================================================================================//
        //                                     Time Zone Standard                                                                              //
        //=====================================================================================================================================//
        private string GetTimeZoneStandard()
        {
            try
            {
                // Get the local time zone
                TimeZoneInfo localTimeZone = TimeZoneInfo.Local;

                // Return the Id of the time zone (e.g., "Pacific Standard Time", "UTC", etc.)
                return localTimeZone.Id;
            }
            catch (Exception ex)
            {
                // Log error and return a default value if something goes wrong
                api.Log(API.LogType.Error, $"Error retrieving time zone: {ex.Message}");
                return "Unknown";
            }
        }



        //=====================================================================================================================================//
        //                                     Day or Night                                                                                    //
        //=====================================================================================================================================//

        private double GetDayOrNight()
        {
            try
            {
                // Get the current system time
                DateTime currentTime = DateTime.Now;

                // Define the start and end times for day (6 AM to 6 PM)
                int startDayHour = 6;
                int endDayHour = 18;

                // Check if the current time is in the day range
                if (currentTime.Hour >= startDayHour && currentTime.Hour < endDayHour)
                {
                    // Day time: Return 1
                    return 1.0;
                }
                else
                {
                    // Night time: Return 2
                    return 2.0;
                }
            }
            catch (Exception ex)
            {
                // Log error if something goes wrong
                api.Log(API.LogType.Error, $"Error determining day/night status: {ex.Message}");
                return 0.0; // Default return in case of error
            }
        }


        //=====================================================================================================================================//
        //                                       Webview2 Install                                                                                //
        //=====================================================================================================================================//


        private bool IsWebViewInstalled()
        {
            try
            {
                string[] registryKeys = new string[]
                {
                    @"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
                    @"SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
                    @"HKEY_CURRENT_USER\Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
                };

                foreach (string registryKey in registryKeys)
                {
                    using (
                        RegistryKey key = registryKey.StartsWith("HKEY_CURRENT_USER")
                            ? Registry.CurrentUser.OpenSubKey(
                                registryKey.Replace("HKEY_CURRENT_USER\\", "")
                            )
                            : Registry.LocalMachine.OpenSubKey(registryKey)
                    )
                    {
                        if (key != null)
                        {
                            Debug($"WebView2 found in registry at: {registryKey}");
                            return true;
                        }
                    }
                }

                Debug("WebView2 not found in registry.");
                return false;
            }
            catch (Exception ex)
            {
                Error($"Error checking WebView2 installation: {ex.Message}");
                return false;
            }
        }

        //=====================================================================================================================================//
        //                                       Chrome Install                                                                                //
        //=====================================================================================================================================//

        private bool IsChromeInstalled()
        {
            try
            {
                string[] possiblePaths = new string[]
                {
                    Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                        "Google",
                        "Chrome",
                        "Application",
                        "chrome.exe"
                    ),
                    Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                        "Google",
                        "Chrome",
                        "Application",
                        "chrome.exe"
                    ),
                    Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "Google",
                        "Chrome",
                        "Application",
                        "chrome.exe"
                    ),
                };

                foreach (string path in possiblePaths)
                {
                    if (File.Exists(path))
                    {
                        Debug($"Chrome found at: {path}");
                        return true;
                    }
                }

                string[] uninstallRegistryKeys = new string[]
                {
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                    @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
                };

                foreach (string registryKey in uninstallRegistryKeys)
                {
                    using (RegistryKey key = Registry.LocalMachine.OpenSubKey(registryKey))
                    {
                        if (key != null)
                        {
                            foreach (string subKeyName in key.GetSubKeyNames())
                            {
                                using (RegistryKey subKey = key.OpenSubKey(subKeyName))
                                {
                                    if (subKey != null)
                                    {
                                        string displayName =
                                            subKey.GetValue("DisplayName") as string;
                                        if (
                                            !string.IsNullOrEmpty(displayName)
                                            && displayName.IndexOf(
                                                "Google Chrome",
                                                StringComparison.OrdinalIgnoreCase
                                            ) >= 0
                                        )
                                        {
                                            Debug("Chrome found in registry.");
                                            return true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                Debug("Chrome not found.");
                return false;
            }
            catch (Exception ex)
            {
                Error($"Error checking for Chrome installation: {ex.Message}");
                return false;
            }
        }

        //=====================================================================================================================================//
        //                                        Internet Connection                                                                          //
        //=====================================================================================================================================//
        private bool IsInternetAvailable()
        {
            try
            {
                Ping ping = new Ping();
                PingReply reply = ping.Send("8.8.8.8", 1000);

                return reply.Status == IPStatus.Success;
            }
            catch (Exception ex)
            {
                Error($"Error checking internet connection: {ex.Message}");
                return false;
            }
        }

        //=====================================================================================================================================//
        //                                         Taskbar Position                                                                            //
        //=====================================================================================================================================//

        public enum TaskbarPosition
        {
            Left = 0,
            Top = 1,
            Right = 2,
            Bottom = 3,
        }

        public class TaskbarUtils
        {
            [DllImport("user32.dll")]
            public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

            [DllImport("user32.dll")]
            public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

            private const string TaskbarClassName = "Shell_TrayWnd";

            public struct RECT
            {
                public int Left;
                public int Top;
                public int Right;
                public int Bottom;
            }

            public static TaskbarPosition GetTaskbarPosition()
            {
                try
                {
                    IntPtr taskbarHandle = FindWindow("Shell_TrayWnd", null);

                    if (taskbarHandle != IntPtr.Zero)
                    {
                        GetWindowRect(taskbarHandle, out RECT rect);

                        int screenWidth = Screen.PrimaryScreen.Bounds.Width;
                        int screenHeight = Screen.PrimaryScreen.Bounds.Height;

                        if (rect.Top == 0 && rect.Bottom != screenHeight)
                        {
                            return TaskbarPosition.Top;
                        }
                        else if (rect.Left == 0 && rect.Right != screenWidth)
                        {
                            return TaskbarPosition.Left;
                        }
                        else if (rect.Right == screenWidth && rect.Left != 0)
                        {
                            return TaskbarPosition.Right;
                        }
                        else
                        {
                            return TaskbarPosition.Bottom;
                        }
                    }
                    else
                    {
                        throw new Exception("Taskbar window not found.");
                    }
                }
                catch (Exception)
                {
                    return TaskbarPosition.Bottom;
                }
            }

            public static string GetTaskbarPositionString()
            {
                TaskbarPosition position = GetTaskbarPosition();

                switch (position)
                {
                    case TaskbarPosition.Left:
                        return "Left";
                    case TaskbarPosition.Top:
                        return "Top";
                    case TaskbarPosition.Right:
                        return "Right";
                    case TaskbarPosition.Bottom:
                        return "Bottom";
                    default:
                        return "Unknown";
                }
            }
        }

        //=====================================================================================================================================//
        //                                         Mouse Y                                                                                     //
        //=====================================================================================================================================//
        private int GetMousePositionY()
        {
            try
            {
                int mouseY = Cursor.Position.Y;
                return mouseY;
            }
            catch (Exception ex)
            {
                Error($"Error retrieving mouse position Y: {ex.Message}");
                return 0;
            }
        }

        //=====================================================================================================================================//
        //                                         Mouse X                                                                                     //
        //=====================================================================================================================================//
        private int GetMousePositionX()
        {
            try
            {
                int mouseX = Cursor.Position.X;
                return mouseX;
            }
            catch (Exception ex)
            {
                Error($"Error retrieving mouse position X: {ex.Message}");
                return 0;
            }
        }

        //=====================================================================================================================================//
        //                                          ScreenHeight                                                                               //
        //=====================================================================================================================================//

        private int GetScreenHeight()
        {
            try
            {
                int screenHeight = Screen.PrimaryScreen.Bounds.Height;
                return screenHeight;
            }
            catch (Exception ex)
            {
                Error($"Error retrieving screen height: {ex.Message}");
                return 0;
            }
        }

        //=====================================================================================================================================//
        //                                          ScreenWidth                                                                                //
        //=====================================================================================================================================//
        private int GetScreenWidth()
        {
            try
            {
                int screenWidth = Screen.PrimaryScreen.Bounds.Width;
                return screenWidth;
            }
            catch (Exception ex)
            {
                Error($"Error retrieving screen width: {ex.Message}");
                return 0;
            }
        }

        //=====================================================================================================================================//
        //                                           IP Adress                                                                                 //
        //=====================================================================================================================================//
        private string GetLocalIPAddress()
        {
            try
            {
                string localIP = string.Empty;
                var host = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName());
                foreach (var ip in host.AddressList)
                {
                    if (
                        ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork
                        && !System.Net.IPAddress.IsLoopback(ip)
                    )
                    {
                        localIP = ip.ToString();
                        break;
                    }
                }

                return !string.IsNullOrEmpty(localIP) ? localIP : "IP Address Not Found";
            }
            catch (Exception ex)
            {
                Error($"Error retrieving IP address: {ex.Message}");
                return "Unknown";
            }
        }

        //=====================================================================================================================================//
        //                                            GenerateStructure                                                                        //
        //=====================================================================================================================================//

        private void CheckAndCopyStructure()
        {
            try
            {
                string nekDataPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "Rainmeter",
                    "NekData"
                );
                string skinResourcePath = Path.Combine(
                    api.ReplaceVariables("#SKINSPATH#"),
                    SkinName,
                    "@Resources",
                    "@Structure"
                );
                string nekDataSkinPath = Path.Combine(nekDataPath, SkinName);
                string versionFilePath = Path.Combine(nekDataSkinPath, $"{SkinVer}.txt");

                if (!Directory.Exists(nekDataPath))
                {
                    Directory.CreateDirectory(nekDataPath);
                    Debug("Created NekData folder at: " + nekDataPath);
                }

                if (!Directory.Exists(skinResourcePath))
                {
                    Error($"Structure folder not found in the @Resources folder of '{SkinName}'.");
                    return;
                }

                if (!Directory.Exists(nekDataSkinPath))
                {
                    Directory.CreateDirectory(nekDataSkinPath);
                    Debug("Created skin folder in NekData: " + nekDataSkinPath);
                }

                string skinVersion = api.ReadString("Version", "").Trim();
                bool copyStructure = true;

                if (File.Exists(versionFilePath))
                {
                    Debug($"Version file found: {versionFilePath}. No need to copy structure.");
                    copyStructure = false;
                }
                else
                {
                    Debug($"Version file not found: {versionFilePath}. Copying structure...");
                }

                if (copyStructure)
                {
                    CopyDirectory(skinResourcePath, nekDataSkinPath);

                    File.WriteAllText(versionFilePath, skinVersion);
                    Debug("Copied Structure folder and created version file: " + versionFilePath);
                }
            }
            catch (Exception ex)
            {
                Error("Error in CheckAndCopyStructure: " + ex.Message);
            }
        }

        //=====================================================================================================================================//
        //                                          ComputerName                                                                               //
        //=====================================================================================================================================//
        private string GetComputerName()
        {
            try
            {
                string computerName = Environment.MachineName;
                return computerName;
            }
            catch (Exception ex)
            {
                Error($"Error retrieving computer name: {ex.Message}");
                return "Unknown";
            }
        }

        //=====================================================================================================================================//
        //                                          Full Window Version                                                                        //
        //=====================================================================================================================================//
        private string GetWindowsFullVersion()
        {
            try
            {
                var version = Environment.OSVersion.Version;
                string fullVersion =
                    $"Version: {version.Major}.{version.Minor}.{version.Build}.{version.Revision}";
                using (
                    var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
                        @"SOFTWARE\Microsoft\Windows NT\CurrentVersion"
                    )
                )
                {
                    if (key != null)
                    {
                        var productName = key.GetValue("ProductName")?.ToString();
                        var releaseId = key.GetValue("ReleaseId")?.ToString();

                        if (!string.IsNullOrEmpty(productName))
                        {
                            fullVersion = $"{productName} {fullVersion}";
                        }

                        if (!string.IsNullOrEmpty(releaseId))
                        {
                            fullVersion += $" (Release ID: {releaseId})";
                        }
                    }
                }

                return fullVersion;
            }
            catch (Exception ex)
            {
                Error($"Error determining full Windows version: {ex.Message}");
                return "Unknown Windows Version";
            }
        }

        //=====================================================================================================================================//
        //                                         UserName                                                                                    //
        //=====================================================================================================================================//
        private string GetUsername()
        {
            try
            {
                string username = Environment.UserName;
                return username;
            }
            catch (Exception ex)
            {
                Error($"Error retrieving username: {ex.Message}");
                return "Unknown";
            }
        }

        //=====================================================================================================================================//
        //                                          OS Bits                                                                                    //
        //=====================================================================================================================================//
        private string GetOSBits()
        {
            try
            {
                bool is64BitOS = Environment.Is64BitOperatingSystem;
                return is64BitOS ? "64-bit" : "32-bit";
            }
            catch (Exception ex)
            {
                Error($"Error determining OS bits: {ex.Message}");
                return "Unknown";
            }
        }

        //=====================================================================================================================================//
        //                                           Windwows Version                                                                          //
        //=====================================================================================================================================//
        private string GetWindowsVersion()
        {
            try
            {
                var version = Environment.OSVersion.Version;
                if (version.Major == 6 && version.Minor == 1)
                    return "Windows 7";
                else if (version.Major == 6 && version.Minor == 2)
                    return "Windows 8";
                else if (version.Major == 6 && version.Minor == 3)
                    return "Windows 8.1";
                else if (version.Major == 10 && version.Minor == 0)
                {
                    return IsWindows11() ? "Windows 11" : "Windows 10";
                }
                else
                    return "Unknown Windows Version";
            }
            catch (Exception ex)
            {
                Error($"Error determining Windows version: {ex.Message}");
                return "Unknown";
            }
        }

        private bool IsWindows11()
        {
            try
            {
                using (
                    var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(
                        @"SOFTWARE\Microsoft\Windows NT\CurrentVersion"
                    )
                )
                {
                    if (key != null)
                    {
                        var productName = key.GetValue("ProductName")?.ToString();
                        if (
                            !string.IsNullOrEmpty(productName) && productName.Contains("Windows 11")
                        )
                        {
                            return true;
                        }
                    }
                }
            }
            catch { }
            return false;
        }

        //=====================================================================================================================================//
        //                                           Total Cores                                                                               //
        //=====================================================================================================================================//
        private int GetTotalPhysicalCores()
        {
            try
            {
                int physicalCoreCount = 0;
                using (
                    var searcher = new System.Management.ManagementObjectSearcher(
                        "SELECT NumberOfCores FROM Win32_Processor"
                    )
                )
                {
                    foreach (var item in searcher.Get())
                    {
                        physicalCoreCount += Convert.ToInt32(item["NumberOfCores"]);
                    }
                }

                return physicalCoreCount;
            }
            catch (Exception ex)
            {
                Error($"Error retrieving total physical cores: {ex.Message}");
                return 0;
            }
        }

        //=====================================================================================================================================//
        //                                         Total  Ram MB                                                                               //
        //=====================================================================================================================================//
        private double GetTotalRamInMB()
        {
            try
            {
                var computerInfo = new Microsoft.VisualBasic.Devices.ComputerInfo();
                double totalRamInBytes = computerInfo.TotalPhysicalMemory;
                double totalRamInMB = totalRamInBytes / (1024.0 * 1024);
                return Math.Round(totalRamInMB, 2);
            }
            catch (Exception ex)
            {
                Error($"Error retrieving total RAM: {ex.Message}");
                return 0.0;
            }
        }

        //=====================================================================================================================================//
        //                                           Ram GB                                                                                    //
        //=====================================================================================================================================//
        private double GetTotalRamInGB()
        {
            try
            {
                var computerInfo = new Microsoft.VisualBasic.Devices.ComputerInfo();
                double totalRamInBytes = computerInfo.TotalPhysicalMemory;
                double totalRamInGB = totalRamInBytes / (1024.0 * 1024 * 1024);
                return Math.Round(totalRamInGB, 2);
            }
            catch (Exception ex)
            {
                Error($"Error retrieving total RAM: {ex.Message}");
                return 0.0;
            }
        }

        //=====================================================================================================================================//
        //                                            GetDirectory Size                                                                        //
        //=====================================================================================================================================//
        private long GetDirectorySize(string folderPath)
        {
            long size = 0;

            foreach (
                string file in Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories)
            )
            {
                FileInfo fileInfo = new FileInfo(file);
                size += fileInfo.Length;
            }

            return size;
        }

        //=====================================================================================================================================//
        //                                            Copy Directory Helper                                                                   //
        //=====================================================================================================================================//
        public void CopyDirectory(string sourceDir, string destDir)
        {
            try
            {
                Directory.CreateDirectory(destDir);

                foreach (var file in Directory.GetFiles(sourceDir))
                {
                    string destFile = Path.Combine(destDir, Path.GetFileName(file));
                    File.Copy(file, destFile, overwrite: true);
                    Debug($"Copied file: {file} to {destFile}");
                }

                foreach (var dir in Directory.GetDirectories(sourceDir))
                {
                    string destSubDir = Path.Combine(destDir, Path.GetFileName(dir));
                    CopyDirectory(dir, destSubDir);
                }
            }
            catch (Exception ex)
            {
                Error($"Error copying directory from {sourceDir} to {destDir}: {ex.Message}");
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
                    string skinResourcePath = Path.Combine(
                        api.ReplaceVariables("#SKINSPATH#"),
                        SkinName,
                        "@Resources"
                    );
                    string patchNoteFilePath = Path.Combine(skinResourcePath, "PatchNote.nek");

                    string patchNoteContent = "[Variables]\nUserName = " + currentUsername;
                    File.WriteAllText(patchNoteFilePath, patchNoteContent);
                    Debug($"Wrote username '{currentUsername}' to PatchNote.nek.");
                    api.Execute(NoMatchAction);
                }
                else
                {
                    Debug("Username matches the system's username. No action required.");
                }
            }
            catch (Exception ex)
            {
                Error("Error in ValidateUserNameAndCopyFonts: " + ex.Message);
            }
        }

        //=====================================================================================================================================//
        //                                            DateModified                                                                             //
        //=====================================================================================================================================//

        private DateTime GetDirectoryLastModified(string folderPath)
        {
            DateTime lastModified = Directory.GetLastWriteTime(folderPath);

            foreach (
                string file in Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories)
            )
            {
                DateTime fileLastModified = File.GetLastWriteTime(file);
                if (fileLastModified > lastModified)
                {
                    lastModified = fileLastModified;
                }
            }

            return lastModified;
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
                Error("Info.dll: 'DeviantArtLinks' is empty or not provided.");
            }

            if (string.IsNullOrEmpty(skinName))
            {
                Error("Info.dll: 'Skin_Name' is empty or not provided.");
            }

            if (string.IsNullOrEmpty(deviantArtLinks) || string.IsNullOrEmpty(skinName))
            {
                return;
            }

            string[] parts = deviantArtLinks.Split(':');
            if (parts.Length != 2)
            {
                Error(
                    "Info.dll: 'DeviantArtLinks' is not formatted correctly. Expected format: link1|link2|link3:Skin1|Skin2|Skin3"
                );
                return;
            }

            string[] linkIds = parts[0].Split('|');
            string[] skinNames = parts[1].Split('|');

            if (linkIds.Length != skinNames.Length)
            {
                Error(
                    "Info.dll: Mismatch between the number of link IDs and skin names in 'DeviantArtLinks'."
                );
                return;
            }

            int index = Array.IndexOf(skinNames, skinName);
            if (index == -1)
            {
                warninig($"Info.dll: Skin name '{skinName}' not found in the skin names list.");
                return;
            }

            string linkId = linkIds[index];
            string deviantArtUrl = $"https://www.deviantart.com/nstechbytes/art/{linkId}";

            Debug($"Info.dll: Constructed DeviantArt URL: {deviantArtUrl}");

            rmbang($"[\"{deviantArtUrl}\"]");
        }

        //=====================================================================================================================================//
        //                                             UnInstall                                                                              //
        //=====================================================================================================================================//

        private void HandleUninstall(string skinsPath, string settingsPath, string skinName)
        {
            string skinFolder = Path.Combine(skinsPath, skinName);
            string docskinFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Rainmeter",
                "NekData",
                skinName
            );
            if (Directory.Exists(skinFolder))
            {
                try
                {
                    Directory.Delete(skinFolder, true);
                    if (Directory.Exists(docskinFolder))
                    {
                        Directory.Delete(docskinFolder, true);
                    }

                    Debug($"Info.dll: Skin folder '{skinFolder}' deleted.");
                }
                catch (Exception ex)
                {
                    Error($"Info.dll: Error deleting skin folder '{skinFolder}': {ex.Message}");
                }
            }
            else
            {
                warninig($"Info.dll: Skin folder '{skinFolder}' does not exist.");
            }

            if (File.Exists(settingsPath))
            {
                try
                {
                    var lines = File.ReadAllLines(settingsPath);
                    var updatedLines = ClearSkinSections(lines, skinName);

                    File.WriteAllLines(settingsPath, updatedLines);
                    Debug($"Info.dll: Removed all sections for '{skinName}' from settings file.");
                }
                catch (Exception ex)
                {
                    Error(
                        $"Info.dll: Error modifying settings file '{settingsPath}': {ex.Message}"
                    );
                }
            }
            else
            {
                warninig($"Info.dll: Settings file '{settingsPath}' does not exist.");
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
                    Debug($"Info.dll: Cleared all sections for '{skinName}' from settings file.");
                }
                catch (Exception ex)
                {
                    Error(
                        $"Info.dll: Error modifying settings file '{settingsPath}': {ex.Message}"
                    );
                }
            }
            else
            {
                Error($"Info.dll: Settings file '{settingsPath}' does not exist.");
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
                        line.Equals($"[{skinName}]", StringComparison.OrdinalIgnoreCase)
                        || line.StartsWith($"[{skinName}\\", StringComparison.OrdinalIgnoreCase);
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
                Error($"Info.dll: Original skin folder '{originalSkinFolder}' does not exist.");
                return;
            }

            if (Directory.Exists(newSkinFolder))
            {
                warninig($"Info.dll: New skin folder '{newSkinFolder}' already exists.");
                return;
            }

            try
            {
                Directory.CreateDirectory(newSkinFolder);

                foreach (
                    string dirPath in Directory.GetDirectories(
                        originalSkinFolder,
                        "*",
                        SearchOption.AllDirectories
                    )
                )
                {
                    Directory.CreateDirectory(dirPath.Replace(originalSkinFolder, newSkinFolder));
                }

                foreach (
                    string filePath in Directory.GetFiles(
                        originalSkinFolder,
                        "*.*",
                        SearchOption.AllDirectories
                    )
                )
                {
                    File.Copy(filePath, filePath.Replace(originalSkinFolder, newSkinFolder), true);
                }

                File.WriteAllText(Path.Combine(newSkinFolder, "Clone.txt"), string.Empty);
                Debug(
                    $"Info.dll: Skin '{originalSkinName}' cloned to '{newSkinName}' with Clone.txt created."
                );
            }
            catch (Exception ex)
            {
                Error($"Info.dll: Error cloning skin: {ex.Message}");
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
