// User preference and administrator policy. Policy always wins over the user's choice.
//
//   Policy (administrators only):  HKLM\SOFTWARE\Policies\Natanor\FocusOverlay   Enabled (DWORD) 0 = off for everyone
//                                  HKCU\SOFTWARE\Policies\Natanor\FocusOverlay   same, per user (Group Policy, user side)
//   User preference:               HKCU\Software\Natanor\FocusOverlay            Mode (string) Always | Search | Off
//                                  (older builds wrote Enabled DWORD 0/1 only; read as Off/Always when Mode is absent)
//
// A policy value of 1 (or no value) leaves the choice to the user. Reads are cheap registry reads, no COM.
using System;
using Microsoft.Win32;

namespace Natanor.FocusOverlay
{
    internal enum PolicyState { NotSet, Allowed, Disabled }

    internal enum FocusMode { Off, Always, Search }

    internal static class Settings
    {
        private const string PolicyPath = @"SOFTWARE\Policies\Natanor\FocusOverlay";
        private const string UserPath = @"Software\Natanor\FocusOverlay";

        public static PolicyState Policy()
        {
            // machine policy first, then user policy: either one set to 0 disables
            var machine = ReadPolicy(Registry.LocalMachine);
            if (machine == PolicyState.Disabled) return PolicyState.Disabled;
            var user = ReadPolicy(Registry.CurrentUser);
            if (user == PolicyState.Disabled) return PolicyState.Disabled;
            return machine == PolicyState.Allowed || user == PolicyState.Allowed ? PolicyState.Allowed : PolicyState.NotSet;
        }

        private static PolicyState ReadPolicy(RegistryKey hive)
        {
            try
            {
                using (var k = hive.OpenSubKey(PolicyPath))
                {
                    if (k == null) return PolicyState.NotSet;
                    object v = k.GetValue("Enabled");
                    if (v == null) return PolicyState.NotSet;
                    return Convert.ToInt32(v) == 0 ? PolicyState.Disabled : PolicyState.Allowed;
                }
            }
            catch { return PolicyState.NotSet; }   // unreadable policy: behave as not set, never crash Excel
        }

        public static FocusMode UserMode
        {
            get
            {
                try
                {
                    using (var k = Registry.CurrentUser.OpenSubKey(UserPath))
                    {
                        if (k == null) return FocusMode.Always;                  // default for a new user
                        string m = k.GetValue("Mode") as string;
                        if (m == "Always") return FocusMode.Always;
                        if (m == "Search") return FocusMode.Search;
                        if (m == "Off") return FocusMode.Off;
                        return Convert.ToInt32(k.GetValue("Enabled", 1)) != 0 ? FocusMode.Always : FocusMode.Off;
                    }
                }
                catch { return FocusMode.Always; }
            }
            set
            {
                try
                {
                    using (var k = Registry.CurrentUser.CreateSubKey(UserPath))
                    {
                        k.SetValue("Mode", value.ToString(), RegistryValueKind.String);
                        k.SetValue("Enabled", value == FocusMode.Off ? 0 : 1, RegistryValueKind.DWord);   // older builds
                    }
                }
                catch (Exception ex) { Log.Write("settings-write-failed " + ex.GetType().Name); }
            }
        }
    }
}
