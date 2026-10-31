// Integration test for the installed add-in (run install.cmd first).
// Creates the add-in through native COM and calls it through real vtables, the
// way Outlook does, so marshalling bugs (like a wrong SAFEARRAY declaration) show up here.
// Exit code 0 = pass.
using System;
using System.Runtime.InteropServices;
using System.Xml;

static class ComTest
{
    const string Clsid = "6F3B5C2A-8E1D-4C7B-9A54-2D7E1F0B3A91";
    const string IidUnknown = "00000000-0000-0000-C000-000000000046";
    const string IidExtensibility = "B65AD801-ABAF-11D0-BB8B-00A0C90F2744";
    const string IidRibbon = "000C0396-0000-0000-C000-000000000046";

    // Vtable slots: 0-2 IUnknown, 3-6 IDispatch, then the interface's own methods.
    const int SlotOnConnection = 7, SlotOnStartupComplete = 10, SlotOnBeginShutdown = 11, SlotGetCustomUI = 7;

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int OnConnectionFn(IntPtr self, IntPtr app, int mode, IntPtr addin,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int OnArrayFn(IntPtr self,
        [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int GetCustomUIFn(IntPtr self, [MarshalAs(UnmanagedType.BStr)] string id,
        [MarshalAs(UnmanagedType.BStr)] out string xml);

    [DllImport("ole32.dll")]
    static extern int CoCreateInstance(ref Guid clsid, IntPtr outer, int ctx, ref Guid iid, out IntPtr obj);

    static int failed;

    static void Check(string name, bool ok, string detail)
    {
        Console.WriteLine((ok ? "ok   " : "FAIL ") + name + (ok || detail == null ? "" : " (" + detail + ")"));
        if (!ok) failed++;
    }

    static T Slot<T>(IntPtr itf, int slot)
    {
        IntPtr vtbl = Marshal.ReadIntPtr(itf);
        return (T)(object)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(vtbl, slot * IntPtr.Size), typeof(T));
    }

    static IntPtr QI(IntPtr unk, string iid)
    {
        IntPtr p;
        Guid g = new Guid(iid);
        int hr = Marshal.QueryInterface(unk, ref g, out p);
        if (hr != 0) throw new Exception("QueryInterface " + iid + " failed: 0x" + hr.ToString("X8"));
        return p;
    }

    static void CheckRibbon(IntPtr rib, string ribbonId)
    {
        string xml;
        int hr = Slot<GetCustomUIFn>(rib, SlotGetCustomUI)(rib, ribbonId, out xml);
        Check("GetCustomUI " + ribbonId, hr == 0 && !string.IsNullOrEmpty(xml), "hr=0x" + hr.ToString("X8"));
        if (string.IsNullOrEmpty(xml)) return;

        var doc = new XmlDocument();
        try { doc.LoadXml(xml); }
        catch (XmlException e) { Check("ribbon XML parses (" + ribbonId + ")", false, e.Message); return; }

        int buttons = doc.SelectNodes("//*[local-name()='group']/*[local-name()='button']").Count;
        int items = doc.SelectNodes("//*[local-name()='contextMenu']//*[local-name()='button']").Count;
        Check("ribbon buttons (" + ribbonId + ")", buttons == 5, buttons + " found");
        Check("context menu items (" + ribbonId + ")", items == 4, items + " found");
    }

    [STAThread]
    static int Main()
    {
        try
        {
            Guid clsid = new Guid(Clsid), iunk = new Guid(IidUnknown);
            IntPtr unk;
            int hr = CoCreateInstance(ref clsid, IntPtr.Zero, 1 /* CLSCTX_INPROC_SERVER */, ref iunk, out unk);
            Check("CoCreateInstance", hr == 0, "hr=0x" + hr.ToString("X8") + "; is the add-in installed?");
            if (hr != 0) return 1;

            IntPtr ext = QI(unk, IidExtensibility);
            Array custom = new object[0];
            Check("OnConnection", Slot<OnConnectionFn>(ext, SlotOnConnection)(ext, unk, 0, unk, ref custom) == 0, null);
            Check("OnStartupComplete", Slot<OnArrayFn>(ext, SlotOnStartupComplete)(ext, ref custom) == 0, null);

            IntPtr rib = QI(unk, IidRibbon);
            CheckRibbon(rib, "Microsoft.Outlook.Mail.Compose");
            CheckRibbon(rib, "Microsoft.Outlook.Explorer");

            Check("OnBeginShutdown", Slot<OnArrayFn>(ext, SlotOnBeginShutdown)(ext, ref custom) == 0, null);
        }
        catch (Exception e)
        {
            Check("no exception", false, e.Message);
        }

        Console.WriteLine(failed == 0 ? "COM integration test passed." : failed + " check(s) failed.");
        return failed == 0 ? 0 : 1;
    }
}
