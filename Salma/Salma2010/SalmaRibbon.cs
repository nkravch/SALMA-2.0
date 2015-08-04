﻿using System.Text.RegularExpressions;
using Office = Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using WordToTFS;

using WordToTFSWordAddIn.Views;

using Microsoft.Office.Interop.Word;

//using SoftwareLocker;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new TfsRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.
namespace Salma2010
{
    

    [ComVisible(true)]
    public class SalmaRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private ThisAddIn addIn;

        private TfsManager tfsManager;

        private string btnConnectLabel;

        private string menuTfsUserLabel;

        private bool userCredentials;

        // Project List
        private int projectsCount;

        // private string projectName = string.Empty;
        private int projectSelectedIndex;

        private Dictionary<int, string> projects = new Dictionary<int, string>();

        // Project Collection List
        private int projectCollectionCount;

        private int projectCollectionSelectedIndex;

        private Dictionary<int, string> projectCollection = new Dictionary<int, string>();

        private bool mnuTfsUserIsEnabled = false;

        private bool cbxConnectionUrlIsEnabled = true;

        private bool btnConnectIsEnabled = true;

        private bool btnNewWorkItemIsEnabled = false;

        private bool btnAddDetailsIsEnabled = false;

        private bool btnLinkItemsIsEnabled = false;

        private bool btnUpdateIsEnabled = false;

        private bool btnUpdateAndSyncIsEnabled = false;

        private bool ddlProjectCollectionIsEnabled = false;

        private bool ddlProjectsIsEnabled = false;

        private bool btnReportIsEnabled = false;

        private bool btnMatrixIsEnabled = false;

        //private bool buttonActivateIsVisible = true;

        private bool buttonShowWorkItemsEnable = false;

        private bool buttonShowWorkItemsVisible = false;

        private bool btnShowWorkItemsIsPressed = false;

        private List<string> ConnectionUrls = new List<string>();

        private int connectionUrlSelectedIndex;

        public string ConnectionUrl { get; set; }

        private bool IsConnected { get; set; }

        #region Ribbon Callbacks

        // Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            RibbonInitControls();
        }

        public void RibbonInitControls()
        {
            addIn = Globals.ThisAddIn;
            ConnectionUrl = string.Empty;
            tfsManager = TfsManager.Instance;
            menuTfsUserLabel = Properties.Resources.lblTextNotLoggedIn;
            btnConnectLabel = Properties.Resources.splitBtnConnectLabel;

            addIn.Application.WindowSelectionChange += TextSelectionChanged;
            IsConnected = false;
            foreach (var url in Properties.Settings.Default.ConnectionURLs)
            {
                ConnectionUrls.Add(url);
            }

            connectionUrlSelectedIndex = 0;

            /*
            if (Properties.Settings.Default.isExpired)
            {
                this.InitTrialDialog();
                btnConnectIsEnabled = false;
                cbxConnectionUrlIsEnabled = false;
            }

            if (Properties.Settings.Default.isFullVerion)
            {
                buttonActivateIsVisible = false;
            }
            */

            if (addIn.MsWordVersion == MsWordVersion.MsWord2013)
            {
                buttonShowWorkItemsVisible = true;
            }
        }

        /// <summary>
        /// GetEnabled
        /// </summary>
        /// <param name="control">
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool EnabledState(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "connectionBox":
                    return cbxConnectionUrlIsEnabled;
                case "menuTfsUser":
                    return mnuTfsUserIsEnabled;
                case "btnConnect":
                    return btnConnectIsEnabled;
                case "ddlProjectCollection":
                    return ddlProjectCollectionIsEnabled;
                case "ddlProjects":
                    return ddlProjectsIsEnabled;
                case "btnNewWorkItem":
                    return btnNewWorkItemIsEnabled;
                case "btnAddDetails":
                    return btnAddDetailsIsEnabled;
                case "btnLinkItems":
                    return btnLinkItemsIsEnabled;

                case "showWorkItems":
                    return buttonShowWorkItemsEnable;

                case "btnUpdate":
                    return btnUpdateIsEnabled;
                case "btnUpdateAndSync":
                    return btnUpdateAndSyncIsEnabled;
                case "btnReport":
                    return btnReportIsEnabled;
                case "btnMatrix":
                    return btnMatrixIsEnabled;
                default:
                    return false;
            }
        }

        public string GetImageMSO(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "btnActivate":
                    return OfficeHelper.GetImageMso(Icons.Activate, addIn.MsWordVersion);
                case "btnHelp_mnu":
                    return OfficeHelper.GetImageMso(Icons.btnHelp, addIn.MsWordVersion);
                case "btnHelp":
                    return OfficeHelper.GetImageMso(Icons.btnHelp, addIn.MsWordVersion);
                case "btnMatrix":
                    return OfficeHelper.GetImageMso(Icons.TraceabilityMatrix, addIn.MsWordVersion);
                case "btnReport":
                    return OfficeHelper.GetImageMso(Icons.Report, addIn.MsWordVersion);
                case "groupHelp":
                    return OfficeHelper.GetImageMso(Icons.Help, addIn.MsWordVersion);
                case "group1":
                    return OfficeHelper.GetImageMso(Icons.Report, addIn.MsWordVersion);
                case "btnUpdate":
                    return OfficeHelper.GetImageMso(Icons.SyncConnectedTool, addIn.MsWordVersion);
                case "btnUpdateAndSync":
                    return OfficeHelper.GetImageMso(Icons.SyncConnectedTool, addIn.MsWordVersion);
                case "showWorkItems":
                    return OfficeHelper.GetImageMso(Icons.ShowCommentsMenu, addIn.MsWordVersion);
                case "btnLinkItems":
                    return OfficeHelper.GetImageMso(Icons.LinkItems, addIn.MsWordVersion);
                case "btnAddDetails":
                    return OfficeHelper.GetImageMso(Icons.AddDetails, addIn.MsWordVersion);
                case "btnNewWorkItem":
                    return OfficeHelper.GetImageMso(Icons.AddNewWorkItem, addIn.MsWordVersion);
                case "groupManageWI":
                    return OfficeHelper.GetImageMso(Icons.groupManageWI, addIn.MsWordVersion);
                case "groupReporting":
                    return OfficeHelper.GetImageMso(Icons.groupReporting, addIn.MsWordVersion);
                case "groupConnect":
                    return OfficeHelper.GetImageMso(Icons.Connect, addIn.MsWordVersion);
                default:
                    return string.Empty;
            }
        }

        public bool VisibleState(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "showWorkItems":
                    return buttonShowWorkItemsVisible;
                default:
                    return false;
            }
        }

        public string GetText(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "connectionBox":
                    {
                        return ConnectionUrl;
                    }

                default:
                    return string.Empty;
            }
        }

        public void OnConnectionUrlChanged(Office.IRibbonControl control, string text)
        {
            ConnectionUrl = text;
        }

        public void OnShowWorkItems(Office.IRibbonControl control, bool IsPressed)
        {
            addIn.Application.CommandBars.ExecuteMso("ShowCommentGutter");
            btnShowWorkItemsIsPressed = addIn.Application.CommandBars.GetPressedMso("ShowCommentGutter");
            ribbon.InvalidateControl("btnShowWorkItems");
        }

        public bool IsPressed(Office.IRibbonControl control)
        {
            return control.Id == "showWorkItems" && btnShowWorkItemsIsPressed;
        }

        public void OnAddDetails(Office.IRibbonControl control)
        {
            addIn.AddDetails();
        }

        /// <summary>
        /// On Add New Work Item
        /// </summary>
        /// <param name="control">
        /// </param>
        public void OnAddNewWorkItem(Office.IRibbonControl control)
        {
            addIn.AddWorkItem(projects[projectSelectedIndex]);
        }

        public string GetButtonImage(Office.IRibbonControl control)
        {
            return IsConnected ? "FileDropSqlDatabase" : "ServerConnection"; // return (IsConnected ? "DisconnectFromServer" : "ConnectToServer");
        }

        public void OnConnectButtonClick(Office.IRibbonControl control)
        {
            if (IsConnected)
            {
                Disconnect();
            }
            else
            {
                if(TfsManager.Instance.Credential != null)
                Connect(TfsManager.Instance.Credential, false);
                else
                    Connect(CredentialCache.DefaultCredentials, false);
            }
        }

        public void UpdateButtonClick(Office.IRibbonControl control)
        {
            addIn.UpdateStatus();
        }
        public void UpdateAndSyncButtonClick(Office.IRibbonControl control)
        {
            if(addIn.MsWordVersion == MsWordVersion.MsWord2013)
                addIn.UpdateStatusAndSync();
            else
                addIn.UpdateStatus();
        }

        public void GenerateReportButtonClick(Office.IRibbonControl control)
        {
            addIn.GenerateReport();
        }

        public void GenerateMatrixButtonClick(Office.IRibbonControl control)
        {
            addIn.GenerateMatrix();
        }

        public void LinkItemsButtonClick(Office.IRibbonControl control)
        {
            addIn.LinkItem();
        }

        /// <summary>
        /// getItemCount
        /// </summary>
        /// <param name="control">
        /// DropDown control
        /// </param>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        public int getItemCount(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "ddlProjects":
                    return projectsCount;
                case "ddlProjectCollection":
                    return projectCollectionCount;
                case "connectionBox":
                    return ConnectionUrls.Count;
                default:
                    return 0;
            }
        }

        /// <summary>
        /// getItemLabel
        /// </summary>
        /// <param name="control">
        /// DropDown control
        /// </param>
        /// <param name="index">
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        public string getItemLabel(Office.IRibbonControl control, int index)
        {
            switch (control.Id)
            {
                case "ddlProjects":
                    return projects[index];
                case "ddlProjectCollection":
                    return projectCollection[index];
                case "connectionBox":
                    return ConnectionUrls[index];
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// getSelectedItemIndex
        /// </summary>
        /// <param name="control">
        /// DropDown control
        /// </param>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        public int getSelectedItemIndex(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "ddlProjects":
                    return projectSelectedIndex;
                case "ddlProjectCollection":
                    return projectCollectionSelectedIndex;
                case "connectionBox":
                    return connectionUrlSelectedIndex;
                default:
                    return 0;
            }
        }

        public void OnAction(Office.IRibbonControl control, string itemID, int itemIndex)
        {
            switch (control.Id)
            {
                case "ddlProjects":
                    {
                        addIn.Project = projects[itemIndex];
                        projectSelectedIndex = itemIndex;

                        TextSelectionChanged(addIn.Application.Selection);
                        return;
                    }

                case "ddlProjectCollection":
                    {
                        projectCollectionSelectedIndex = itemIndex;
                        tfsManager.ChangeCollection(projectCollection[itemIndex]);
                        addIn.TeamProjectCollectionName = projectCollection[itemIndex];
                        PopulateProjects();
                        ribbon.InvalidateControl("ddlProjects");
                        return;
                    }

                case "connectionBox":
                    {
                        connectionUrlSelectedIndex = itemIndex;
                        return;
                    }

                default:
                    return;
            }
        }

        public string GetConnectLabel(Office.IRibbonControl control)
        {
            return btnConnectLabel;
        }

        public string GetLabelText(Office.IRibbonControl control)
        {
            if (control.Id == "menuTfsUser")
            {
                return menuTfsUserLabel;
            }

            string controlId = control.Id;

            if ("btnConnect" == control.Id && IsConnected)
            {
                controlId = "btnDisconnect";
            }

            if ("connectionBox" == control.Id)
            {
                controlId = "lblConnectionUrl";
            }

            PropertyInfo[] properties = typeof(Properties.Resources).GetProperties(BindingFlags.Static | BindingFlags.NonPublic);
            PropertyInfo prop = properties.Select(p => p).Where(p => p.Name == controlId + "Label").FirstOrDefault();

            string label = string.Empty;

            if (prop != null)
            {
                label = (string)prop.GetValue(null, null);
            }

            return label;
        }

        /*
        public bool IsVisible(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "btnActivate":
                    return buttonActivateIsVisible;
                default:
                    return false;
            }
        }

        public void OnActivateButtonClick(Office.IRibbonControl control)
        {
            if (InitTrialDialog() == false)
            {
                Properties.Settings.Default.isFullVerion = true;
                Properties.Settings.Default.Save();
                buttonActivateIsVisible = false;
                ribbon.InvalidateControl("btnActivate");
            }
        }*/

        public void OnHelpButtonClick(Office.IRibbonControl control)
        {
            CultureInfo ci = new CultureInfo((int)addIn.Application.Language);
            if (ci.Name == "ru-RU")
            {
                Help.ShowHelp(null, "Help/How_to_Use_SALMA_RU.chm");
            }
            else
            {
                Help.ShowHelp(null, "Help/How_to_Use_SALMA.chm");
            }
        }

        /*
        private bool? InitTrialDialog()
        {
            try
            {
                var location = Assembly.GetExecutingAssembly().Location;
                string path = Path.GetDirectoryName(location);
                var trialMaker = new TrialMaker("WordToTFS", path + "\\RegFile.reg", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\SALMA\\TMSetp.dbf", String.Empty, 30, "745");

                byte[] tdKey = { 97, 250, 1, 5, 84, 21, 7, 63, 4, 54, 87, 56, 123, 10, 3, 62, 7, 9, 20, 36, 37, 21, 101, 57 };
                trialMaker.TripleDESKey = tdKey;

                TrialMaker.RunTypes runTypes = trialMaker.ShowDialog(Properties.Resources.btnActivateLabel, OleInterop.GetMsoImage(Icons.Activate, 32, 32));
                if (runTypes != TrialMaker.RunTypes.Expired)
                {
                    return runTypes != TrialMaker.RunTypes.Full;
                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show(ex.ToString());
            }
            return null;
        }
        */

        private void TextSelectionChanged(Selection selection)
        {
            bool shapeSelected = selection.ShapeRange.Count != 0;

            bool textSelected = ((!String.IsNullOrWhiteSpace(selection.Range.Text)) || shapeSelected) && IsConnected;

            btnAddDetailsIsEnabled = textSelected;
            btnNewWorkItemIsEnabled = textSelected;
            ribbon.InvalidateControl("btnNewWorkItem");
            ribbon.InvalidateControl("btnAddDetails");

            if (selection.Comments.Count > 0)
            {
                var commentText = addIn.GetCommentTextBySelection();

                btnLinkItemsIsEnabled = !string.IsNullOrWhiteSpace(commentText) && projects[projectSelectedIndex] == addIn.ParseProjectName(commentText);
            }
            else
            {
                btnLinkItemsIsEnabled = false;
            }

            ribbon.InvalidateControl("btnLinkItems");
        }

        #endregion

        #region Helpers

        /// <summary>
        /// Connect to TFS
        /// </summary>
        private void Connect(ICredentials cred, bool anotherUser)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(ConnectionUrl))
                {
                    MessageBox.Show(ResourceHelper.GetResourceString("MSG_CONNECTION_URL_IS_EMPTY"), ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    IsConnected = tfsManager.Connect(new Uri(ConnectionUrl), cred, anotherUser);
                    AfterState(IsConnected);
                }
            }
            catch (UriFormatException ex)
            {
                MessageBox.Show(ResourceHelper.GetResourceString("MSG_CONNECTION_URL_IS_INCORRECT"), ResourceHelper.GetResourceString("ERROR_TEXT"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                EventLog.WriteEntry("SalmaConnection", ex.Message, EventLogEntryType.Error);
            }
        }

        private void Disconnect()
        {
            
            IsConnected = false;
            AfterState(IsConnected);
        }

        /// <summary>
        /// Populate Project Collection
        /// </summary>
        private void PopulateProjectCollection()
        {
            projectCollection.Clear();
            var collection = tfsManager.GetTeamProjectCollection();
            for (var i = 0; i < collection.Count; i++)
            {
                projectCollection.Add(i, collection[i]);
            }

            projectCollectionCount = collection.Count;
            projectCollectionSelectedIndex = 0;
            addIn.TeamProjectCollectionName = projectCollection[0];
        }

        /// <summary>
        /// Populate Projects
        /// </summary>
        private void PopulateProjects()
        {
            projects.Clear();
            var projectList = tfsManager.GetProjectsForCurrentTPC(projectCollectionSelectedIndex);
            projectList.Sort();
            for (var i = 0; i < projectList.Count(); i++)
            {
                projects.Add(i, projectList[i]);
            }

            projectsCount = projects.Count;
            projectSelectedIndex = 0;
            addIn.Project = projects[projectSelectedIndex];
            
        }

        private void AfterState(bool isConnected)
        {
            if (isConnected)
            {
                PopulateProjectCollection();
                PopulateProjects();
                btnConnectLabel = Properties.Resources.btnDisconnectLabel;
                menuTfsUserLabel = tfsManager.GetUserDisplayName();
                StoreConnectionUrl();
                
            }
            else
            {
                // TODO:
                projects.Clear();
                projects.Add(0, string.Empty);
                projectsCount = 0;
                projectSelectedIndex = 0;
                projectCollection.Clear();
                projectCollection.Add(0, string.Empty);
                projectCollectionCount = 0;
                projectCollectionSelectedIndex = 0;
                btnConnectLabel = Properties.Resources.btnConnectLabel;
                menuTfsUserLabel = Properties.Resources.lblTextNotLoggedIn;
            }

            cbxConnectionUrlIsEnabled = !isConnected;
            ddlProjectCollectionIsEnabled = isConnected;
            ddlProjectsIsEnabled = isConnected;

            btnUpdateIsEnabled = isConnected;
            btnUpdateAndSyncIsEnabled = isConnected;
            mnuTfsUserIsEnabled = isConnected;
            btnNewWorkItemIsEnabled = false;
            btnAddDetailsIsEnabled = false;
            btnLinkItemsIsEnabled = false;
            btnMatrixIsEnabled = isConnected;
            btnReportIsEnabled = isConnected;

            buttonShowWorkItemsEnable = isConnected;
            ribbon.Invalidate();
        }

        /// <summary>
        /// Store last five successful connections in app settings
        /// </summary>
        private void StoreConnectionUrl()
        {
            string clearUrl = RemoveExtraSlashes(ConnectionUrl);
            if (!Properties.Settings.Default.ConnectionURLs.Contains(clearUrl))
            {
                ConnectionUrls.Insert(0, clearUrl);
                Properties.Settings.Default.ConnectionURLs.Insert(0, clearUrl);
                if (Properties.Settings.Default.ConnectionURLs.Count > 4)
                {
                    Properties.Settings.Default.ConnectionURLs.RemoveAt(
                        Properties.Settings.Default.ConnectionURLs.Count - 1);
                }
                Properties.Settings.Default.Save();
            }
        }

        private string RemoveExtraSlashes(string url)
        {
            string protocol = url.Contains("http://") ? "http://" : "https://";
            url = url.Replace(protocol, string.Empty);
            url = Regex.Replace(url, "//+", "/");
            if (url[url.Length-1] == '/')
                url=url.TrimEnd('/');
            url=url.Insert(0, protocol);
          return url;

        }

        public void ConnectAsAnotherUser(Office.IRibbonControl control)
        {
            tfsManager.Disconnect();
            Disconnect();    
            Connect(new NetworkCredential(), true);
        }


        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (var resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }

            return null;
        }

        #endregion

        public string getSupertip(Office.IRibbonControl control)
        {
            string controlId = control.Id;

            if ("btnConnect" == control.Id && IsConnected)
            {
                controlId = "btnConnect_disconnect";
            }

            var properties = typeof(Properties.Resources).GetProperties(BindingFlags.Static | BindingFlags.NonPublic);
            var prop = properties.Select(p => p).FirstOrDefault(p => p.Name == controlId + "ToolTip");

            string label = string.Empty;

            if (prop != null)
            {
                label = (string)prop.GetValue(null, null);
            }

            return label;
        }

        public string getScreentip(Office.IRibbonControl control)
        {
            return GetLabelText(control);
        }

        public void InitAttributes(string element, string attribute, XmlDocument doc)
        {
            var list = doc.GetElementsByTagName(element);

            foreach (XmlNode node in list)
            {
                var attr = doc.CreateAttribute(attribute);
                attr.Value = attribute;
                node.Attributes.Append(attr);
            }
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            string xml = string.Empty;
            if (Globals.ThisAddIn.MsWordVersion == MsWordVersion.MsWord2007)
            {
                xml = GetResourceText("Salma2010.SalmaRibbon2007.xml");
            }
            else
            {
                xml = GetResourceText("Salma2010.SalmaRibbon.xml");
            }

            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);

            InitAttributes("button", "getScreentip", doc);
            InitAttributes("comboBox", "getScreentip", doc);
            InitAttributes("dropDown", "getScreentip", doc);
            InitAttributes("toggleButton", "getScreentip", doc);

            InitAttributes("button", "getSupertip", doc);
            InitAttributes("comboBox", "getSupertip", doc);
            InitAttributes("dropDown", "getSupertip", doc);
            InitAttributes("toggleButton", "getSupertip", doc);
            InitAttributes("menu", "getSupertip", doc);

            return doc.OuterXml;
        }

        #endregion
    }
}