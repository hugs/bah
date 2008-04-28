using System;
using SHDocVw;
using MSHTML;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Expando;
using System.Reflection;


namespace BHOBrowserAutomationHelper
{
    public static class Win32
	{
		[DllImport("user32.dll", SetLastError = true)]
		public static extern bool SetForegroundWindow(int hWnd);
	}

    [
    ComVisible(true),
    ClassInterface(ClassInterfaceType.None)
    ]
    public class ScriptableObject : IScriptableObject
    {
        WebBrowser webBrowser;
        HTMLDocument document; 

        public ScriptableObject(WebBrowser browser)
        {
            webBrowser = browser;
            document = webBrowser.Document as HTMLDocument;
        }

        public void helloWorld()
        {
            System.Windows.Forms.MessageBox.Show("Hello, World!");
        }

        public void showMessage(String textToShow)
        {
            System.Windows.Forms.MessageBox.Show(textToShow);
        }

        public void setFileField(String fieldId, String filePath)
        {
            IHTMLInputElement fileInputField = (IHTMLInputElement)document.getElementById(fieldId);
            // TODO - Do more checking to verify SetForegroundWindow worked, 
            //        and complain if it doesn't.
            Win32.SetForegroundWindow(webBrowser.HWND);
            fileInputField.select();
            System.Windows.Forms.SendKeys.SendWait(filePath);
        }

    }

    [
    ComVisible(true),
    // Replace the following GUID with your own.
    // Get one at guidgen.com.
    Guid("8a194578-81ea-4850-9911-13ba2d71efbd"),
    ClassInterface(ClassInterfaceType.None)
    ]
    public class BHO : IObjectWithSite
    {
        WebBrowser webBrowser;
        HTMLDocument document;
        IHTMLWindow2 window;
        IExpando winExpando;

        public void OnDocumentComplete(object pDisp, ref object URL)
        {
            document = webBrowser.Document as HTMLDocument;
            window = document.parentWindow as IHTMLWindow2;

            // In the next few lines, we're going to add a callable
            // object to the "window" namespace exposed to the
            // JavaScript in the browser.
            winExpando = window as IExpando;
            IScriptableObject so = new ScriptableObject(webBrowser);

            // Warning!!! You (yes, *you*!) should change "bah" to a unique string for your project.
            PropertyInfo myProp = winExpando.AddProperty("bah");
            myProp.SetValue(winExpando, so, null);
        }

        #region BHO Internal Functions

        public static string BHOKEYNAME = "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects";

        [ComRegisterFunction]
        public static void RegisterBHO(Type type)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(BHOKEYNAME, true);

            if (registryKey == null)
                registryKey = Registry.LocalMachine.CreateSubKey(BHOKEYNAME);

            string guid = type.GUID.ToString("B");
            RegistryKey ourKey = registryKey.OpenSubKey(guid);

            if (ourKey == null)
                ourKey = registryKey.CreateSubKey(guid);

            ourKey.SetValue("Alright", 1);
            registryKey.Close();
            ourKey.Close();
        }

        [ComUnregisterFunction]
        public static void UnregisterBHO(Type type)
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(BHOKEYNAME, true);
            string guid = type.GUID.ToString("B");

            if (registryKey != null)
                registryKey.DeleteSubKey(guid, false);
        }

        public int SetSite(object site)
        {
            if (site != null)
            {
                webBrowser = (WebBrowser)site;
                webBrowser.DocumentComplete += new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
            }
            else
            {
                webBrowser.DocumentComplete -= new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
                webBrowser = null;
            }

            return 0;
        }

        public int GetSite(ref Guid guid, out IntPtr ppvSite)
        {
            IntPtr punk = Marshal.GetIUnknownForObject(webBrowser);
            int hr = Marshal.QueryInterface(punk, ref guid, out ppvSite);
            Marshal.Release(punk);

            return hr;
        }

        #endregion
    }
}
