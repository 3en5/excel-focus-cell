// POC-01/02 host: a plain COM add-in (IDTExtensibility2), compiled with the .NET Framework csc (C# 5).
// No VSTO dependency at this stage; ARCH-01 decides the final host after the POC.
using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Natanor.FocusOverlay
{
    // Same IID and vtable order as Extensibility.IDTExtensibility2; declared here to avoid shipping Extensibility.dll
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), TypeLibType(4160)]   // FDual | FDispatchable
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application, [In] int connectMode,
                                      [In, MarshalAs(UnmanagedType.IDispatch)] object addInInst,
                                      [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection([In] int removeMode,
                                         [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    // AutoDispatch: Office calls the ribbon callbacks by name through IDispatch on this object
    [ComVisible(true), Guid("6E1D7B0A-3C52-4F8E-9A41-2B7C5D0F1E93"), ProgId("NatanorFocusOverlay.Connect"),
     ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, Office.IRibbonExtensibility
    {
        private Controller _controller;
        private Office.IRibbonUI _ribbon;
        private bool _he = true;   // ribbon language follows Excel's UI language: Hebrew (1037) or English

        private string T(string he, string en) { return _he ? he : en; }

        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            Log.Write("on-connection mode=" + connectMode);
            try
            {
                try { _he = ((Excel.Application)application).LanguageSettings.LanguageID[Office.MsoAppLanguageID.msoLanguageIDUI] == 1037; }
                catch { _he = true; }
                _controller = new Controller((Excel.Application)application);
                _controller.StateChanged += () => { if (_ribbon != null) _ribbon.Invalidate(); };
                Log.Write("connected");
            }
            catch (Exception ex) { Log.Write("connect-failed " + ex.GetType().Name + " 0x" + ex.HResult.ToString("X8")); }
        }

        public void OnDisconnection(int removeMode, ref Array custom)
        {
            _ribbon = null;
            if (_controller != null) { _controller.Dispose(); _controller = null; }
            // Excel waits for every COM reference to be released before its process can exit. Our RCWs (Application,
            // event sinks, ranges) are released only when the GC runs; without this, an Excel closed through
            // automation (Quit) stayed alive with no window (seen 21/10/2026). Collect our own garbage only - no blind
            // FinalReleaseComObject on objects other components may share.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Log.Write("disconnected mode=" + removeMode);
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // ---- Ribbon: group "מיקוד תא" at the end of the View tab with two mutually exclusive toggles ---------------
        //   nfoAlways  "תמיד"     - highlight all the time
        //   nfoSearch  "בחיפוש"  - highlight only while the Find and Replace dialog is open
        // Pressing the pressed one again switches the highlight off.

        public string GetCustomUI(string ribbonId)
        {
            string xml = @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnRibbonLoad"">
  <ribbon><tabs><tab idMso=""TabView"">
    <group id=""nfoGroup"" label=""{G}"">
      <toggleButton id=""nfoAlways"" size=""large"" label=""{A}"" getImage=""GetImage""
        getPressed=""GetPressed"" onAction=""OnToggle"" getEnabled=""GetEnabled""
        getScreentip=""GetScreentip"" getSupertip=""GetSupertip""/>
      <toggleButton id=""nfoSearch"" size=""large"" label=""{S}"" getImage=""GetImage""
        getPressed=""GetPressed"" onAction=""OnToggle"" getEnabled=""GetEnabled""
        getScreentip=""GetScreentip"" getSupertip=""GetSupertip""/>
    </group>
  </tab></tabs></ribbon>
</customUI>";
            return xml.Replace("{G}", T("מיקוד תא", "Focus")).Replace("{A}", T("תמיד", "Always")).Replace("{S}", T("בחיפוש", "On search"));
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbon) { _ribbon = ribbon; }

        public object GetImage(Office.IRibbonControl control) { return RibbonIcon.Picture(32, control.Id == "nfoSearch"); }

        private FocusMode ModeOf(Office.IRibbonControl control)
        {
            return control.Id == "nfoSearch" ? FocusMode.Search : FocusMode.Always;
        }

        public bool GetPressed(Office.IRibbonControl control)
        {
            return _controller != null && _controller.Mode == ModeOf(control);
        }

        public void OnToggle(Office.IRibbonControl control, bool pressed)
        {
            if (_controller == null) return;
            _controller.SetMode(pressed ? ModeOf(control) : FocusMode.Off);
            if (_ribbon != null) _ribbon.Invalidate();      // the other button must release
        }

        public bool GetEnabled(Office.IRibbonControl control)
        {
            return _controller != null && !_controller.LockedByPolicy;
        }

        public string GetScreentip(Office.IRibbonControl control)
        {
            return ModeOf(control) == FocusMode.Search ? T("מיקוד תא בזמן חיפוש", "Focus while searching") : T("מיקוד תא תמיד", "Focus always");
        }

        public string GetSupertip(Office.IRibbonControl control)
        {
            if (_controller != null && _controller.LockedByPolicy) return T("הושבת על ידי מנהל המערכת.", "Disabled by your administrator.");
            if (ModeOf(control) == FocusMode.Search)
                return T("מדגיש את השורה והעמודה של התא הפעיל רק כשחלון \"חיפוש והחלפה\" (Ctrl+F) פתוח. לחיצה נוספת מכבה. ההדגשה היא תצוגה בלבד ואינה משנה את הקובץ.",
                         "Highlights the active cell's row and column only while the Find and Replace dialog (Ctrl+F) is open. Click again to turn off. The highlight is display-only and never changes the file.");
            return T("מדגיש את השורה והעמודה של התא הפעיל כל הזמן. לחיצה נוספת מכבה. ההדגשה היא תצוגה בלבד ואינה משנה את הקובץ.",
                     "Highlights the active cell's row and column all the time. Click again to turn off. The highlight is display-only and never changes the file.");
        }
    }
}
