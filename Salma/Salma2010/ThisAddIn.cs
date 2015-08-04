using System;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using WordToTFS;
using WordToTFS.Model;
using WordToTFS.ViewModel.CreateNew;
using WordToTFSWordAddIn.Views;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Net.Mime;
using System.Security.Cryptography;
using System.Windows.Controls.Primitives;
using System.Windows.Media.Media3D;
using Microsoft.SqlServer.Server;
using WordToTFS.ConfigHelpers;

namespace Salma2010
{
    public partial class ThisAddIn
    {
        public string Project { get; set; }
        public string TeamProjectCollectionName { get; set; }
        protected int WorkItemId
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the ms word version.
        /// </summary>
        public MsWordVersion MsWordVersion { get; private set; }


        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            var app = this.GetHostItem<Microsoft.Office.Interop.Word.Application>(typeof(Microsoft.Office.Interop.Word.Application), "Application");
            var ci = new CultureInfo((int)app.Language);
            Thread.CurrentThread.CurrentUICulture = ci;

            MsWordVersion = OfficeHelper.GetMsWordVersion(app.Version);
            SectionManager.SetSection(MsWordVersion);
            return new SalmaRibbon();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            WindowFactory.IconConverter = OleInterop.GetMsoImage;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion

        public void AddWorkItem(string projectName)
        {
            var ListBoxCollection = new ListBoxItems(Application.Selection.Text,TfsManager.Instance.GetWorkItemsTypeForCurrentProject(projectName));
            var popup = new CreateNew();
            popup.DataContext = ListBoxCollection;
            popup.Create(null, Icons.AddNewWorkItem);
            if (popup.isCancelled && !popup.isCreated)
                return;
            
            var type = ListBoxCollection.GetValue();
            var title = ListBoxCollection.GetTitle();
            var workItemId = CreateNewWi.AddWorkItemForCurrentProject(projectName, title, type);
            if (workItemId != 0)
            {
                var wi = TfsManager.Instance.GetWorkItemText(workItemId);
                object text = string.Format("{0} {1}", SalmaConstants.Comments.WI_ID, workItemId);
                var comment = Application.ActiveDocument.Comments.Add(Application.Selection.Range, ref text);
                comment.Range.Hyperlinks.Add(comment.Range, TfsManager.Instance.GetWorkItemLink(workItemId, Project));
                comment.Range.InsertAfter(wi);
                WorkItemId = workItemId;
            }

        }

        /// <summary>
        /// Adding description
        /// </summary>
        public void AddDetails()
        {
            var popup = new AddDetails();
            var lowerConnectionUrl = TfsManager.Instance.GetTfsUrl().ToLower();
            var lowerCollectionName = TeamProjectCollectionName.ToLower();
            var wiNumbers = new List<int>();
            List<string> itemsSource = new List<string>();
            //List<string> itemsSource2 = new List<string>();
            foreach (Comment comment in Application.ActiveDocument.Comments)
            {
                var commentLink = comment.Range.Hyperlinks.Cast<Hyperlink>().FirstOrDefault();
                if(commentLink != null)
                {
                    var url = commentLink.Address.ToLower();
                    if (url.Contains(lowerConnectionUrl) && url.Contains(lowerCollectionName))
                    {
                        int number = TfsManager.ParseId(comment.Range.Text);
                        if (!comment.Range.Text.Contains(SalmaConstants.Comments.FOR_WI_ID) && !wiNumbers.Contains(number))
                        {
                            itemsSource.Add(comment.Range.Text);
                            //itemsSource2.Add(comment.Range.Text);
                            wiNumbers.Add(number);
                        }
                    }
                }
            }

            popup.ListBox.SelectedItem = 2;
            popup.FoundItemsListView.SelectedItem = popup.FoundItemsListView.AlternationCount;
            popup.ListBox.Items.Clear();
            if (WorkItemId != 0)
            {
                for (int i = 0; i < itemsSource.Count; i++)
                {
                    if (itemsSource[i].Contains(WorkItemId.ToString()))
                    {
                        popup.ListBox.ScrollIntoView(itemsSource[i]);
                        popup.ListBox.SelectedIndex = i;
                    }
                }
            }
            popup.ListBox.Focus();
            popup.ListBox.ItemsSource = itemsSource;
            var types = TfsManager.Instance.GetWorkItemsTypeForCurrentProject(Project);
            types.Add(ResourceHelper.GetResourceString("ALL"));

            popup.FilterBox.ItemsSource = types.OrderBy(f => f).ToList();

            popup.FilterBox.SelectedValue = ResourceHelper.GetResourceString("ALL");

            popup.FilterBox.SelectionChanged += (s, a) =>
            { popup.ListBox.ItemsSource = ((string)a.AddedItems[0] == ResourceHelper.GetResourceString("ALL")) ? itemsSource : itemsSource.Where(item => item.Contains(string.Format("{0}{1}", ResourceHelper.GetResourceString("TYPE"), a.AddedItems[0]))).ToList(); };

            popup.ListBox.SelectionChanged += (s, a) =>
                {
                    if (popup.ListBox.SelectedValue != null)
                    {
                        popup.AddDetailsAsBox.IsEnabled = true;
                    }
                    else
                    {
                        popup.AddDetailsAsBox.IsEnabled = false;
                        popup.AddDetailsAsBox.ItemsSource = null;
                    }
                };

            popup.Create(null, Icons.AddDetails);

            if (popup.IsAdd || popup.IsReplace)
            {
                string wi = string.Empty;
                int id = 0;
                //if current document TabItem tab taking data from this tab
                if (popup.CurrentDocumentTabItem())
                {
                    wi = (string)popup.ListBox.SelectedValue;
                    id = ParseId(wi);
                }
                //if By Work Item ID TabItem tab taking data from this tab
                if (popup.ByWorkItemIDTabItem())
                {
                    id = Convert.ToInt32(popup.GetWIID.Text);
                }
                wi = TfsManager.Instance.GetWorkItemText(id);
                wi = wi.Replace("\n", "\r");

                //Application.Selection.Copy();
                CopySelectionToClipboard();

                bool com = true;
                if (Clipboard.ContainsData(DataFormats.Html))
                {
                    //TODO: Fix clipboard access bug
                    IDataObject dataObject = null;
                    for (var i = 0; i < 10; i++)
                    {
                        try
                        {
                            dataObject = Clipboard.GetDataObject();
                        }
                        catch
                        {
                            //And you thought this would be easy.....
                        }
                        break;
                    }

                    
                    if (dataObject != null)
                    {
                        string description = TfsManager.Instance.GetWorkItemDescription(id, popup.AddDetailsAsBox.SelectedValue.ToString());

                        if (popup.IsAdd && description == string.Empty)
                        {
                            TfsManager.Instance.AddDetailsForWorkItem(id, popup.AddDetailsAsBox.SelectedValue.ToString(), dataObject, out com);
                        }
                        else
                        {
                            TfsManager.Instance.ReplaceDetailsForWorkItem(id, popup.AddDetailsAsBox.SelectedValue.ToString(), dataObject,out com);
                        }
                    }
                }
                else if (Clipboard.ContainsData(DataFormats.Bitmap))
                {
                    var bitmapSource = Clipboard.GetImage();
                    using (var ms = new MemoryStream())
                    {
                        bitmapSource.Save(ms, ImageFormat.Png);

                        string description = TfsManager.Instance.GetWorkItemDescription(id, popup.AddDetailsAsBox.SelectedValue.ToString());

                        if (popup.IsAdd && description == string.Empty)
                        {
                            TfsManager.Instance.AddDetailsForWorkItem(id, popup.AddDetailsAsBox.SelectedValue.ToString(), ms);
                        }
                        else
                        {
                            TfsManager.Instance.ReplaceDetailsForWorkItem(id, popup.AddDetailsAsBox.SelectedValue.ToString(), ms);
                        }
                    }
                }

                if (com)
                {
                    DeleteOldComment(id, popup.AddDetailsAsBox.SelectedValue.ToString());
                    object text = string.Format("{0} {1}", SalmaConstants.Comments.WI_ID, id);
                    var comment = Application.ActiveDocument.Comments.Add(Application.Selection.Range, ref text);
                    comment.Range.Hyperlinks.Add(comment.Range, TfsManager.Instance.GetWorkItemLink(id, Project));
                    comment.Range.InsertBefore(string.Format("{0} {1} ", popup.AddDetailsAsBox.SelectedValue.ToString(), ResourceHelper.GetResourceString("FOR")));
                    comment.Range.InsertAfter(String.Format(wi));
                }
            }
        }

        /// <summary>
        ///Delete old comment for detail
        /// </summary>
        private void DeleteOldComment(int id, string fieldName)
        {
            foreach (Comment comment in Application.ActiveDocument.Comments)
            {
                string commentText = comment.Range.Text;

                if (commentText.Contains(string.Format(" {0} {1}", SalmaConstants.Comments.FOR_WI_ID, id)))
                {
                    int commentTextFor = commentText.IndexOf(SalmaConstants.Comments.FOR);
                    string commentFieldName = commentText.Substring(0, commentTextFor - 1);

                    if (commentFieldName != String.Empty && commentFieldName.ToLower().Equals(fieldName.ToLower()))
                    {
                        if (this.MsWordVersion == MsWordVersion.MsWord2013)
                            comment.DeleteRecursively();
                        else
                            comment.Delete();
                    }
                }
            }
        }

        /// <summary>
        /// Copy selection to clipboard
        /// </summary>
        private void CopySelectionToClipboard()
        {
            // Get last elem in range
            Range lastCharacter = Application.Selection.Range.Characters.Last;

            // If it's text selection
            if (Application.Selection.Type == WdSelectionType.wdSelectionNormal && lastCharacter != null)
            {
                // If last paragraph mark isn't closed - close it
                if (!(lastCharacter.Text.Equals(Convert.ToChar(0x000D).ToString()) ||
                lastCharacter.Text.Equals(Convert.ToChar(0x000D).ToString() + Convert.ToChar(0x0007).ToString())))
                {
                    Range selectionRange = Application.Selection.Document.Range(Application.Selection.Range.Start, 
                        Application.Selection.Range.End);
                    selectionRange.InsertParagraphAfter();
                    selectionRange.Copy();
                    selectionRange.SetRange(selectionRange.End - 1, selectionRange.End);
                    selectionRange.Delete();
                    return;
                }
            }

            //copy to clipboard
            Application.Selection.Copy();
        }

        public void UpdateStatus()
        {
            var lowerConnectionUrl = TfsManager.Instance.GetTfsUrl().ToLower();
            var lowerCollectionName = TeamProjectCollectionName.ToLower();
            Comments comments;
            
            foreach (Comment comment in Application.ActiveDocument.Comments)
            {
                var commentLink = comment.Range.Hyperlinks.Cast<Hyperlink>().FirstOrDefault();
                if (commentLink != null)
                {
                    var url = commentLink.Address.ToLower();
                    if (url.Contains(lowerConnectionUrl) && url.Contains(lowerCollectionName))
                    {
                        
                        var commentText = comment.Range.Text;
                        comment.Range.Text = string.Empty;
                        
                        comment.Range.SetRange(0, 0);
                        var id = ParseId(commentText);
                        int linksCount;
                        
                        var wi = TfsManager.Instance.GetWorkItemText(id, out linksCount);
                        var text = string.Format("{0} {1}", SalmaConstants.Comments.WI_ID, id);
                        comment.Range.InsertBefore(text);
                        comment.Range.Hyperlinks.Add(comment.Range, TfsManager.Instance.GetWorkItemLink(id, Project));
                        if (commentText.Contains(string.Format(" {0} {1}", SalmaConstants.Comments.FOR_WI_ID, id)))
                        
                        { 

                            comment.Range.InsertBefore(
                                commentText.Substring(0, commentText.IndexOf(":") + 2).TrimStart());
                        }

                        comment.Range.InsertAfter(wi);
                        
                        if (linksCount > 0) comment.Range.InsertAfter("\nLinks: " + linksCount);
                                                                           
                        }
                        
                    }
                }
            
        }

        // 
        // Update Title and Description in Document from WI if it was changed inside TFS (changed by user, computed etc.). Also sync "State" like in UpdateStatus() method.
        //
        // Idea: Check documment comments one by one, if comment is "Salma comment" - check if contains Title or Description(Any HTML WI Field).
        // If "YES" compare date of comment creation and date of this field Changed. If Field Changed - delete previous comment with document text, paste new one and create new comment automatically
        // If Field not changed - get status of current WI in tfs, Rewrite text to comment.
        //
        //
        //
        //

        public void UpdateStatusAndSync() //Added for MS Marketing Event. Checks all comments in document, if Title or Description(Another HTML Field) changed in TFS WI - updating text and comment in document.
        {
            //if (Application.ActiveDocument.TrackRevisions == true) // If in current document TrackChanges ON
            //{
                //string message ="SALMA is going to accept all tracking changes in this document. Do you wish to proceed?";
                //string caption = "SALMA – Accept All Changes required";
                // Show message box

                System.Reflection.PropertyInfo[] properties = typeof(Properties.Resources).GetProperties(System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);

                string message = string.Empty;
                string caption = string.Empty;

                System.Reflection.PropertyInfo prop = properties.Select(p => p).Where(p => p.Name == "lblAcceptAllChangesRequiredLabel").FirstOrDefault();

                if (prop != null)
                    message = (string)prop.GetValue(null, null);

                prop = properties.Select(p => p).Where(p => p.Name == "lblAcceptAllChangesRequiredCaptionLabel").FirstOrDefault();

                if (prop != null)
                    caption = (string)prop.GetValue(null, null);

                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                DialogResult result;
                result = MessageBox.Show(message, caption, buttons);


                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    Application.ActiveDocument.AcceptAllRevisions();
                    Application.ActiveDocument.TrackRevisions = false;
                }
                else
                {
                    return;
                }
            //}     
               
            #region Process Comments and Update Text if Changed

            var lowerConnectionUrl = TfsManager.Instance.GetTfsUrl().ToLower();
            var lowerCollectionName = TeamProjectCollectionName.ToLower();
            Application.ActiveDocument.TrackRevisions = true; // Turn ON Track Changes


            Application.ActiveWindow.View.RevisionsFilter.Markup = WdRevisionsMarkup.wdRevisionsMarkupSimple;

            var aa = Application.ActiveWindow.ActivePane;

            //get comments to update
            List<Comment> commentsToUpdate = (from comment in Application.ActiveDocument.Comments.OfType<Comment>()
                                              select comment).ToList<Comment>();

            //foreach (Comment comment in Application.ActiveDocument.Comments)
            foreach (Comment comment in commentsToUpdate)
            {
                var commentLink = comment.Range.Hyperlinks.Cast<Hyperlink>().FirstOrDefault();
                if (commentLink != null)
                {
                    var url = commentLink.Address.ToLower();
                    if (url.Contains(lowerConnectionUrl) && url.Contains(lowerCollectionName)) // Check If comment has link to WI
                    {
                        var commentText = comment.Range.Text; //Comment.Range.Text - is a text in comment Section (like in comment bubble)
                        comment.Range.SetRange(0, 0);
                        var id = ParseId(commentText);
                        int linksCount;
                        var wi = TfsManager.Instance.GetWorkItemText(id, out linksCount); //Get necessary text from WI fields
                            
                        if (commentText.Contains(string.Format(" {0} {1}", SalmaConstants.Comments.FOR_WI_ID, id))) //If comment has link for any Description or HTML field
                        {
                                
                            int commentTextFor = commentText.IndexOf(SalmaConstants.Comments.FOR);  // Get Field Name From comment range text
                            string fieldName = commentText.Substring(0, commentTextFor - 1);        //

                            if (fieldName != String.Empty) 
                            {
                                bool historyDescription = TfsManager.Instance.GetWorkItemFieldHistory(id,fieldName,comment.Date); // Check if Text Changed in WI Revisions
                                if(historyDescription == true)
                                {
                                        Range oldCommentScopeRange;
                                        string description = TfsManager.Instance.GetWorkItemDescription(id,fieldName);

                                        /*
                                            CopyToClipboard(description);

                                        
                                            if (comment.Scope.Start == 0)
                                            {
                                                oldCommentScopeRange = Application.ActiveDocument.Range(0, comment.Scope.End + 1); //Select Range of Current Comment if comment is first line on page
                                            }
                                            else
                                            {
                                                oldCommentScopeRange = Application.ActiveDocument.Range(comment.Scope.Start, comment.Scope.End + 1); // selectrange of current comment
                                            }

                                            comment.Scope.Text = "";
                                            oldCommentScopeRange.Paragraphs.Add().Range.Paste(); // Paste new paragraph to Selected Range
                                            */

                                        oldCommentScopeRange = PasteFromClipboard(comment, description);

                                        Comment newcomment = Application.ActiveDocument.Comments.Add(oldCommentScopeRange); //Creating New comment for pasted range

                                        //Fill comment additional information
 
                                        var text = string.Format("{0} {1}", SalmaConstants.Comments.WI_ID, id);
                                        newcomment.Range.InsertBefore(text);
                                        newcomment.Range.Hyperlinks.Add(newcomment.Range,
                                            TfsManager.Instance.GetWorkItemLink(id, Project));

                                        newcomment.Range.InsertBefore(commentText.Substring(0, commentText.IndexOf(":") + 2).TrimStart());

                                        newcomment.Range.InsertAfter(wi);

                                        if (linksCount > 0) comment.Range.InsertAfter("\nLinks: " + linksCount);
                                        //Fill comment additional text     
                                }
                                else // Just Update Current comment status
                                {
                                    comment.Range.Text = string.Empty;
                                    var text = string.Format("{0} {1}", SalmaConstants.Comments.WI_ID, id);
                                    comment.Range.InsertBefore(text);
                                    comment.Range.Hyperlinks.Add(comment.Range,
                                        TfsManager.Instance.GetWorkItemLink(id, Project));
                                    if (
                                        commentText.Contains(string.Format(" {0} {1}",
                                            SalmaConstants.Comments.FOR_WI_ID, id)))
                                    {

                                        comment.Range.InsertBefore(
                                            commentText.Substring(0, commentText.IndexOf(":") + 2).TrimStart());
                                    }

                                    comment.Range.InsertAfter(wi);

                                    if (linksCount > 0) comment.Range.InsertAfter("\nLinks: " + linksCount);
                                }
                            }
                        }
                        else // If current comment link to WI and contains WI Title
                        {
                            bool historyTitle = TfsManager.Instance.GetWorkItemFieldHistory(id, "Title",comment.Date); //Check if Title was changed in revisions
                            if (historyTitle == true)
                            {
                                string wiTitle = TfsManager.Instance.GetWorkItemTitle(id);

                                CopyToClipboard(wiTitle); //Copy Title
                                Style style = comment.Scope.get_Style();
                                Font font = comment.Scope.Font.Duplicate;
                                Range rng = Application.ActiveDocument.Range(comment.Scope.Start,comment.Scope.End); //Set new comment range
                                Range oldCommentScopeRange;

                                if (comment.Scope.Start == 0)
                                {
                                    oldCommentScopeRange = Application.ActiveDocument.Range(0,comment.Scope.End + 1); //if comment first line on page
                                }
                                else
                                {
                                    oldCommentScopeRange = Application.ActiveDocument.Range(comment.Scope.Start,comment.Scope.End + 1); 
                                }

                                //Fill comment additional information
                                oldCommentScopeRange.Paragraphs.Add().Range.PasteSpecial(WdPasteOptions.wdMatchDestinationFormatting); //insert new paragpaph
                              

                                Comment newcomment = oldCommentScopeRange.Comments.Add(oldCommentScopeRange); // , oldCommentScopeRange.Paragraphs[1].Range.Text); //Create new Comment
                                  
                                //
                                var text = string.Format("{0} {1}", SalmaConstants.Comments.WI_ID, id);
                                newcomment.Range.InsertBefore(text);
                                newcomment.Range.Hyperlinks.Add(newcomment.Range,
                                    TfsManager.Instance.GetWorkItemLink(id, Project));

                                newcomment.Range.InsertAfter(wi);

                                if (linksCount > 0) newcomment.Range.InsertAfter("\nLinks: " + linksCount);
                                //Fill comment additional information

                            }
                            else
                            {
                                //just Update Current comment status
                                    commentText = comment.Range.Text;
                                comment.Range.Text = string.Empty;

                                comment.Range.SetRange(0, 0);
                                id = ParseId(commentText);
                                   
                                wi = TfsManager.Instance.GetWorkItemText(id, out linksCount);
                                var text = string.Format("{0} {1}", SalmaConstants.Comments.WI_ID, id);
                                comment.Range.InsertBefore(text);
                                comment.Range.Hyperlinks.Add(comment.Range, TfsManager.Instance.GetWorkItemLink(id, Project));
                                comment.Range.InsertAfter(wi);

                                if (linksCount > 0) comment.Range.InsertAfter("\nLinks: " + linksCount);

                                //just Update Current comment status
                            }
                        }
                    }
                }

            } 
        #endregion
         
        }

        /// <summary>
        /// Paste selection from clipboard
        /// </summary>
        private Range PasteFromClipboard(Comment comment, string description)
        {
            string r = Convert.ToChar(0x000D).ToString(); // "\r"
            string ra = Convert.ToChar(0x000D).ToString() + Convert.ToChar(0x0007).ToString(); // "\r\a"

            // get comment ranges
            Range commentRange = Application.ActiveDocument.Range(comment.Scope.Start, comment.Scope.End);

            //expand comment to the end of table
            Paragraph p = commentRange.Paragraphs.Last;

            while (p.Range.Text.EndsWith(ra))
            {
                p = p.Next(1);
            }

            // get new range
            commentRange.SetRange(comment.Scope.Start, p.Range.End);

            // delete range
            commentRange.Delete();

            // copy and paste from clipboard
            CopyToClipboard(description);
            commentRange.Paste();

            // delete old comment
            if (this.MsWordVersion == MsWordVersion.MsWord2013)
                comment.DeleteRecursively();
            else
                comment.Delete();

            return commentRange;
        }

        //public void UpdateHistory()
        //{
        //    foreach (Comment comment in Application.ActiveDocument.Comments)
        //    {
        //        var commentText = comment.Range.Text;
        //        if (!commentText.Contains(SalmaConstants.Comments.FOR_WI_ID))
        //        {
        //            var id = ParseId(commentText);

        //            var collection = TfsManager.Instance.GetWorkItemHistory(id);
        //            comment.Range.Bold = 1;
        //            comment.Range.InsertAfter("\n-------------------------------------");
        //            foreach (var item in collection)
        //            {
        //                comment.Range.InsertAfter("\n" + item);
        //            }
        //        }
        //    }
        //}

        public void AddHistory()
        {
            foreach (Comment comment in Application.ActiveDocument.Comments)
            {
                var commentText = comment.Range.Text;
                var id = ParseId(commentText);
                TfsManager.Instance.AddHistoryToWorkItem(id, commentText);
            }
        }

        private static int ParseId(string text)
        {
            string[] splText = text.Split('\r');
            string temp = null;

            for (int i = 0; i < splText.Length; i++)
            {
                if (splText[i].IndexOf(ResourceHelper.GetResourceString("WI_ID")) != -1)
                {
                    temp = splText[i].Split(':').LastOrDefault();
                    break;
                }
            }

            return Int32.Parse(temp.Trim());
        }

        internal String ParseProjectName(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            //Project: SSCC
            //Links: 1
            var temp = text.Substring(text.IndexOf(ResourceHelper.GetResourceString("PROJECT"), StringComparison.Ordinal));
            if (temp.Contains("\r"))
            {


                temp = temp.Substring(ResourceHelper.GetResourceString("PROJECT").Length);
                temp = temp.Substring(0, temp.IndexOf("\r"));
            }
            else
            {
                temp = temp.Substring(ResourceHelper.GetResourceString("PROJECT").Length);
            }
            return temp;
        }

        internal void SetCurrentProject(string projectName)
        {
            Project = projectName;
        }

        internal string GetCommentTextBySelection()
        {
            Comment comment;
            if (String.IsNullOrWhiteSpace(Application.Selection.Text))
            {
                comment = Application.Selection.Comments.Cast<Comment>().FirstOrDefault(c => Application.Selection.Range.Start >= c.Range.Start && Application.Selection.Range.Start <= c.Range.End);

            }
            else
            {
                var text = Application.Selection.Text.Trim();
                comment = Application.Selection.Comments.Cast<Comment>().FirstOrDefault(c => c.Scope.Text != null && c.Scope.Text.Trim() == text) ?? Application.Selection.Comments.Cast<Comment>().FirstOrDefault(c => c.Range.Text != null && c.Range.Text.Trim() == text);
            }

            if (comment != null)
            {
                return comment.Range.Text;
            }
            return string.Empty;
        }

        /// <summary>
        /// Link Work Item 
        /// </summary>
        public void LinkItem()
        {
            var text = GetCommentTextBySelection();
            if (!string.IsNullOrWhiteSpace(text))
            {
                var id = ParseId(text);


                var itemsSource = Application.ActiveDocument.Comments.Cast<Comment>()
                         .Select(comment => comment.Range.Text)
                         .Where(commentText => commentText != null &&
                             !commentText.Contains(SalmaConstants.Comments.FOR_WI_ID) &&
                             commentText.Contains(SalmaConstants.Comments.WI_ID) &&
                                !commentText.Contains(string.Format("{0} {1}", SalmaConstants.Comments.WI_ID, id)) &&
                                commentText.Contains(Project)).ToList();

                //if (TfsManager.Instance.LinkItem(itemsSource, id, project))
                LinkWorkItem LinkWorkItem = new LinkWorkItem(Application.ActiveDocument);
                if (LinkWorkItem.LinkItem(itemsSource, id, Project))
                {
                    UpdateStatus();
                }

            }
        }

        public void GenerateReport()
        {
            var report = new Report(Application.ActiveDocument);
            report.GenerateReport(Project, NormalText);
        }

        public void GenerateMatrix()
        {
            var matrix = new TraceabilityMatrix(Application.ActiveDocument);
            matrix.GenerateMatrix(Project);
        }

        public static void CopyToClipboard(string html)
        {
            var builder = new StringBuilder();
            builder.Append("Version:0.9\r\nStartHTML:{0:000000}\r\nEndHTML:{1:000000}\r\nStartFragment:{2:000000}\r\nEndFragment:{3:000000}\r\n");
            builder.AppendFormat("<html>\r\n<head>\r\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset={0}\">\r\n<title>HTML clipboard</title>\r\n</head>\r\n<body>\r\n<!--StartFragment-->\r\n", Encoding.UTF8.WebName);
            builder.Append(html);
            builder.Append("<!--EndFragment-->\r\n</body>\r\n</html>\r\n");

            //DataFormats.GetFormat(DataFormats.Html);
            DataObject obj = new DataObject();
            obj.SetData(DataFormats.Html, builder.ToString());
            Clipboard.SetDataObject(obj, true);
        }

        private void SetClipboard(object text)
        {
            try
            {
                CopyToClipboard((string)text);
            }
            catch (Exception)
            {
            }
        }
        private void NormalText(string text)
        {
            if (String.IsNullOrWhiteSpace(text))
            {
                return;
            }

            //StringBuilder sb = new StringBuilder();
            //sb.Append("<html>");
            //sb.Append("<body>");
            //sb.Append(text);
            //sb.Append("</body>");
            //sb.Append("</html>");
            //text = sb.ToString();

            Object styleTitle = WdBuiltinStyle.wdStyleNormalObject;
            Object styleTitle2 = WdBuiltinStyle.wdStyleHtmlNormal;
            var doc = Application.ActiveDocument;
            var pt = OfficeHelper.CreateParagraphRange(ref doc);
            pt.Text = text;
            pt.set_Style(ref styleTitle2);

            //text = text.Replace("&lt;", "<");
            //text = text.Replace("&gt;", ">");
            //text = text.Replace("&amp;", "");
            //text = text.Replace("nbsp;", "");

            var thread = new Thread(SetClipboard);
            if (thread.TrySetApartmentState(ApartmentState.STA))
            {
                thread.Start(text);
                thread.Join();
            }

            try
            {
                object objDataTypeMetafile = WdPasteDataType.wdPasteHTML;
                //object objDataTypeMetafile = WdPasteDataType.wdPasteText;
                pt.PasteSpecial(DataType: objDataTypeMetafile);
                if (Application.ActiveDocument.InlineShapes.Count > 0)
                {
                    var page = Application.ActiveDocument.PageSetup;
                    float calculatedWidth = page.PageWidth - (page.LeftMargin + page.RightMargin);

                    foreach (InlineShape shape in Application.ActiveDocument.InlineShapes)
                    {
                        shape.LockAspectRatio = MsoTriState.msoTrue;

                        if (shape.Width <= calculatedWidth) continue;

                        shape.Width = calculatedWidth;
                    }
                }
            }
            catch (Exception)
            {
            }

            Application.ActiveDocument.Content.InsertParagraphAfter();
        }

        private void NormalTextSyncDescription(string text, Comment comment)
        {
            if (String.IsNullOrWhiteSpace(text))
            {
                return;
            }

            //StringBuilder sb = new StringBuilder();
            //sb.Append("<html>");
            //sb.Append("<body>");
            //sb.Append(text);
            //sb.Append("</body>");
            //sb.Append("</html>");
            //text = sb.ToString();

            Object styleTitle = WdBuiltinStyle.wdStyleNormalObject;
            Object styleTitle2 = WdBuiltinStyle.wdStyleHtmlNormal;
            var doc = Application.ActiveDocument;
            var pt = OfficeHelper.CreateParagraphRange(ref doc);
            pt.Text = text;
            pt.set_Style(ref styleTitle2);

            //text = text.Replace("&lt;", "<");
            //text = text.Replace("&gt;", ">");
            //text = text.Replace("&amp;", "");
            //text = text.Replace("nbsp;", "");

            var thread = new Thread(SetClipboard);
            if (thread.TrySetApartmentState(ApartmentState.STA))
            {
                thread.Start(text);
                thread.Join();
            }

            try
            {
                object objDataTypeMetafile = WdPasteDataType.wdPasteHTML;
                //object objDataTypeMetafile = WdPasteDataType.wdPasteText;
                pt.PasteSpecial(DataType: objDataTypeMetafile);
                if (Application.ActiveDocument.InlineShapes.Count > 0)
                {
                    var page = Application.ActiveDocument.PageSetup;
                    float calculatedWidth = page.PageWidth - (page.LeftMargin + page.RightMargin);

                    foreach (InlineShape shape in Application.ActiveDocument.InlineShapes)
                    {
                        shape.LockAspectRatio = MsoTriState.msoTrue;

                        if (shape.Width <= calculatedWidth) continue;

                        shape.Width = calculatedWidth;
                    }
                }
            }
            catch (Exception)
            {
            }
            
            //comment.Scope.InsertParagraph();
            Application.ActiveDocument.Content.InsertParagraphAfter();
        }
    }



}
