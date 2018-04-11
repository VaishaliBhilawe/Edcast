using SHDocVw;
using mshtml;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System;

namespace GuideMePlayer
{
    /// <summary>
    /// Set the GUID of this class and specify that this class is ComVisible.
    /// A BHO must implement the interface IObjectWithSite. 
    /// </summary>
    [ComVisible(true),
    ClassInterface(ClassInterfaceType.None),
    Guid("C42D40F0-BEBF-418D-8EA1-18D99AC2AB17")]

    
    public class BHO : IObjectWithSite
    {
        private InternetExplorer ieInstance;
        public const string BHO_REGISTRY_KEY_NAME = "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects";
        private const string EXTENSIONNAME = "GuideMePlayer";
        private const string JSPATH = "//guidemeprod.blob.core.windows.net/guideme-player-ie/guideme.js";
        private const string USERKEY = "FD5B8ED4-31D8-4015-8E87-4D8C5A0B0B45";

        private WebBrowser webBrowser;

        #region Com Register/UnRegister Methods
        /// <summary>
        /// When this class is registered to COM, add a new key to the BHORegistryKey 
        /// to make IE use this BHO.
        /// On 64bit machine, if the platform of this assembly and the installer is x86,
        /// 32 bit IE can use this BHO. If the platform of this assembly and the installer
        /// is x64, 64 bit IE can use this BHO.
        /// </summary>
        [ComRegisterFunction]
        public static void RegisterBHO(Type t)
        {


            
            RegistryKey key = Registry.LocalMachine.OpenSubKey(BHO_REGISTRY_KEY_NAME, true);
            if (key == null)
            {
                key = Registry.LocalMachine.CreateSubKey(BHO_REGISTRY_KEY_NAME);
            }

            
            // 32 digits separated by hyphens, enclosed in braces: 
            // {00000000-0000-0000-0000-000000000000}
            string bhoKeyStr = t.GUID.ToString("B");

            RegistryKey bhoKey = key.OpenSubKey(bhoKeyStr, true);
            
            // Create a new key.
            if (bhoKey == null)
            {
                bhoKey = key.CreateSubKey(bhoKeyStr);
            }

            // NoExplorer:dword = 1 prevents the BHO to be loaded by Explorer
            string name = "NoExplorer";
            object value = (object)1;
            bhoKey.SetValue(name, value);
            key.Close();
            bhoKey.Close();

            RegistryKey classkey = Registry.ClassesRoot.OpenSubKey("CLSID\\" + bhoKeyStr, true);
            if (classkey == null)
            {
                classkey = Registry.ClassesRoot.CreateSubKey("CLSID\\" + bhoKeyStr);
            }

            classkey.CreateSubKey("Control");
            classkey.CreateSubKey("Implemented Categories\\{59fb2056-d625-48d0-a944-1a85b5ab2640}");
            classkey.CreateSubKey("MiscStatus").SetValue("", "0");
            classkey.CreateSubKey("TypeLib").SetValue("",Marshal.GetTypeLibGuidForAssembly(t.Assembly).ToString("B"));
            classkey.CreateSubKey("Programmable");
            
            string version = t.Assembly.GetName().Version.Major.ToString() + "." + t.Assembly.GetName().Version.Minor.ToString();
            if (version == "0.0") version = "1.0";

            classkey.CreateSubKey("Version").SetValue("", version);


            classkey.Close();

        }

        /// <summary>
        /// When this class is unregistered from COM, delete the key.
        /// </summary>
        [ComUnregisterFunction]
        public static void UnregisterBHO(Type t)
        {
            RegistryKey key = Registry.LocalMachine.OpenSubKey(BHO_REGISTRY_KEY_NAME, true);
            string guidString = t.GUID.ToString("B");
            if (key != null)
            {
                key.DeleteSubKey(guidString, false);
            }

            

            try
            {
                Registry.ClassesRoot.DeleteSubKeyTree(
                "CLSID\\" + t.GUID.ToString("B"));
            }
            catch (ArgumentException) { }
        }
        #endregion

        #region IObjectWithSite Members
        /// <summary>
        /// This method is called when the BHO is instantiated and when
        /// it is destroyed. The site is an object implemented the 
        /// interface InternetExplorer.
        /// </summary>
        /// <param name="site"></param>
        public void SetSite(Object site)
        {
            if (site != null)
            {
                ieInstance = (InternetExplorer)site;
                webBrowser = (WebBrowser) site;

                // Register the DocumentComplete event. 
                webBrowser.DocumentComplete += new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
                webBrowser.DownloadComplete += new DWebBrowserEvents2_DownloadCompleteEventHandler(this.OnDownloadComplete);
            }
        }

        /// <summary>
        /// Retrieves and returns the specified interface from the last site
        /// set through SetSite(). The typical implementation will query the
        /// previously stored pUnkSite pointer for the specified interface.
        /// </summary>
        public void GetSite(ref Guid guid, out Object ppvSite)
        {
            IntPtr punk = Marshal.GetIUnknownForObject(ieInstance);
            ppvSite = new object();
            IntPtr ppvSiteIntPtr = Marshal.GetIUnknownForObject(ppvSite);
            int hr = Marshal.QueryInterface(punk, ref guid, out ppvSiteIntPtr);
            Marshal.ThrowExceptionForHR(hr);
            Marshal.Release(punk);
            Marshal.Release(ppvSiteIntPtr);
        }
        #endregion

        #region event handler



        private void OnDownloadComplete()
        {
            HTMLDocument document = (HTMLDocument) webBrowser.Document;
            this.InjectJavascript(document);
        }

        public void OnDocumentComplete(object pDisp, ref object URL)
        {
            HTMLDocument document = (HTMLDocument) webBrowser.Document;
            this.InjectJavascript(document);
        }

        public void InjectJavascript(HTMLDocument document)
        {

            string script = "window.guideMe = {userKey:'" + USERKEY + "', trackingId:''};";
            string url = JSPATH;
            if (document != null)
            {

                mshtml.IHTMLElementCollection headElementCollection = document.getElementsByTagName("head");
                if (headElementCollection != null)
                {
                    mshtml.IHTMLElement injectedScript = document.getElementById("__guideme_script");
                    if (injectedScript == null)
                    {

                        mshtml.IHTMLElement headElement = headElementCollection.item(0, 0) as mshtml.IHTMLElement;
                        mshtml.IHTMLElement guideme = (mshtml.IHTMLElement)document.createElement("script");
                        mshtml.IHTMLScriptElement guidemeScript = (mshtml.IHTMLScriptElement)guideme;
                        guidemeScript.text = script;

                        mshtml.IHTMLDOMNode guidemeNode = (mshtml.IHTMLDOMNode)guideme;
                        guideme.id = "__guideme_script_config";

                        mshtml.IHTMLElement element = (mshtml.IHTMLElement)document.createElement("script");
                        mshtml.IHTMLScriptElement scriptElement = (mshtml.IHTMLScriptElement)element;
                        scriptElement.src = url;

                        element.id = "__guideme_script";
                        mshtml.IHTMLDOMNode node = (mshtml.IHTMLDOMNode)element;
                        mshtml.HTMLBody body = document.body as mshtml.HTMLBody;
                        body.appendChild(guidemeNode);
                        body.appendChild(node);
                    }

                    Marshal.ReleaseComObject(document);
                }
            }
        }
        #endregion


    }
}
