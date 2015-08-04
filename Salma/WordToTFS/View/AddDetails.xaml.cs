using System;
using System.Windows;
using System.Windows.Controls;
using WordToTFS;
using System.Collections.Generic;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Threading;
using System.Windows.Media;
using MessageBox = System.Windows.MessageBox;

namespace WordToTFSWordAddIn.Views
{
    public partial class AddDetails : Window
    {
        public ListBox ListBox { set { listBox1 = value; } get { return listBox1; } }
        public ComboBox FilterBox { set { filterBox = value; } get { return filterBox; } }
        public ComboBox AddDetailsAsBox { set { comboAddDetailsAs = value; } get { return comboAddDetailsAs; } }
        public TextBox GetWIID { set { workItemID = value; } get { return workItemID; } }
        public ListBox FoundItemsListView { set { foundItemsListView = value; } get { return foundItemsListView; } }
        public TextBox WorkItemID { set { workItemID = value; } get { return workItemID; } }
        private string externalWiIdsText = string.Empty;
        bool correctValueInserted = false;
        public List<WorkItem> WorkItemsToLink = new List<WorkItem>();

        private readonly Timer keyEntryTimer;

        public AddDetails()
        {
            InitializeComponent();
            
            keyEntryTimer = new Timer(UpdateItems, null, -1, -1);
        }

        public bool IsCanceled { set; get; }
        public bool IsAdd { set; get; }
        public bool IsReplace { set; get; }

        public bool IsEmpty { get; set; }

        private void AddButtonClick(object sender, RoutedEventArgs e)
        {
            this.IsCanceled = false;
            IsAdd = true;
            IsReplace = false;
            Close();
        }

        private void ReplaceButtonClick(object sender, RoutedEventArgs e)
        {
            IsReplace = true;
            IsCanceled = false;
            IsAdd = false;
            Close();
        }

        public bool ByWorkItemIDTabItem()
        {
            return byWorkItemIDTabItem.IsSelected;
        }

        public bool CurrentDocumentTabItem()
        {
            return currentDocumentTabItem.IsSelected;
        }

        private void CancelButtonClick(object sender, RoutedEventArgs e)
        {
            IsCanceled = true;
            IsReplace = false;
            IsAdd = false;
            Close();
        }

        public void UpdateItems(object state)
        {
            Dispatcher.Invoke((Action)delegate()
            {
                WorkItemsToLink.Clear();
                foundItemsListView.Items.Clear();
            });

            int itemId = 0;

            if (string.IsNullOrEmpty(externalWiIdsText))
            {
                IsEmpty = true;
            }
            else
            {
                IsEmpty = false;
            }

            if (!String.IsNullOrWhiteSpace(externalWiIdsText))
            {
                int id;
                if (int.TryParse(externalWiIdsText, out id))
                {
                    itemId = id;
                }
                else
                {
                    Dispatcher.BeginInvoke((Action)delegate()
                    {
                        foundItemsListView.Items.Add(new ListViewItem()
                        {
                            Content = String.Format(ResourceHelper.GetResourceString("MSG_INPUT_VALUE_INCORRECT"), externalWiIdsText),
                            Background = new SolidColorBrush(Colors.LightCoral)
                        });
                        SetAddReplaceButton(false, false);
                        AddDetailsAsBox.IsEnabled = false;
                        AddDetailsAsBox.ItemsSource = null;
                        correctValueInserted = false;
                    });
                }
            }

            WorkItem wItem = TfsManager.Instance.GetWorkItem(itemId);
            if (wItem != null)
            {
                Dispatcher.BeginInvoke((Action)delegate()
                {
                    foundItemsListView.Items.Add(String.Format("• {0} {1} ({2}): {3} ", wItem.Type.Name, wItem.Id, wItem.State, wItem.Title));
                    WorkItemsToLink.Add(wItem);
                });

            }
            else
            {
                if (!IsEmpty)
                {
                    Dispatcher.BeginInvoke((Action)delegate()
                    {
                        int id;
                        if (int.TryParse(externalWiIdsText, out id))
                        {
                            foundItemsListView.Items.Add(new ListViewItem()
                            {
                                Content = String.Format(ResourceHelper.GetResourceString("MSG_ITEM_IS_NOT_FOUND"), itemId),
                                Background = new SolidColorBrush(Colors.LightCoral)
                            });
                        }
                    });
                }
            }
        }

        private void workItemID_TextChanged(object sender, TextChangedEventArgs e)
        {
            externalWiIdsText = workItemID.Text;

            if (string.IsNullOrEmpty(externalWiIdsText))
                IsEmpty = true;
            
            AddDetailsAsBox.IsEnabled = true;
            int wi;
            //int.TryParse(workItemID.Text, out wi);

            if (int.TryParse(workItemID.Text, out wi))
            {
                AddDetailsAsBox.ItemsSource = TfsManager.Instance.GetHtmlFieldsByItemId(wi);
                var defaultfield = TfsManager.Instance.GetDefaultDetailsFieldName(wi);
                AddDetailsAsBox.SelectedValue = defaultfield;

                if (AddDetailsAsBox.SelectedValue != null)
                {
                    string description = TfsManager.Instance.GetWorkItemDescription(wi, defaultfield);

                    if (description == string.Empty)
                        SetAddReplaceButton(true, false);
                    else
                        SetAddReplaceButton(false, true);

                    correctValueInserted = true;
                    keyEntryTimer.Change(500, -1);
                    return;
                    /*
                    AddDetailsAsBox.SelectedValue = defaultfield;
                    if (AddDetailsAsBox.SelectedValue == "Steps" || AddDetailsAsBox.SelectedValue == "Шаги")
                    {
                        replaceButton.IsEnabled = false;
                    }
                    addButton.IsEnabled = true;
                    replaceButton.IsEnabled = true;
                    correctValueInserted = true;*/
                }
            }
          
            SetAddReplaceButton(false, false);
            AddDetailsAsBox.IsEnabled = false;
            AddDetailsAsBox.ItemsSource = null;
            correctValueInserted = false;
          
            keyEntryTimer.Change(500, -1); 
        }

        private void SetAddReplaceButton(bool add, bool replace)
        {
            addButton.IsEnabled = add;
            replaceButton.IsEnabled = replace;
        }

        private void listBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SetAddReplaceButton(true, true);
        }

        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tabControl = sender as TabControl;

            if (tabControl != null)
            {
                switch (tabControl.SelectedIndex)
                {
                    case 0:
                        if (ListBox.SelectedValue != null)
                        {
                            var id = TfsManager.ParseId(ListBox.SelectedValue.ToString());
                            var defaultField = TfsManager.Instance.GetDefaultDetailsFieldName(id);
                            AddDetailsAsBox.ItemsSource = TfsManager.Instance.GetHtmlFieldsByItemId(id);
                            AddDetailsAsBox.SelectedValue = defaultField;

                            if (AddDetailsAsBox.SelectedValue != null)
                            {
                                string description = TfsManager.Instance.GetWorkItemDescription(id, defaultField);

                                if (description == string.Empty)
                                    SetAddReplaceButton(true, false);
                                else
                                    SetAddReplaceButton(false, true);

                                return;

                                /*
                                if (AddDetailsAsBox.SelectedValue.ToString() == "Steps" || comboAddDetailsAs.Text == "Шаги")
                                {
                                    replaceButton.IsEnabled = false;
                                }*/
                            }
                        }

                        break;
                    case 1:
                        if (correctValueInserted)
                        {
                            int id;
                            if (int.TryParse(workItemID.Text, out id))
                            {
                                var defaultField = TfsManager.Instance.GetDefaultDetailsFieldName(id);
                                AddDetailsAsBox.ItemsSource = TfsManager.Instance.GetHtmlFieldsByItemId(id);
                                AddDetailsAsBox.SelectedValue = defaultField;

                                if (AddDetailsAsBox.SelectedValue != null)
                                {
                                    string description = TfsManager.Instance.GetWorkItemDescription(id, defaultField);

                                    if (description == string.Empty)
                                        SetAddReplaceButton(true, false);
                                    else
                                        SetAddReplaceButton(false, true);

                                    return;
                                }
                            }
                        }
                        break;
                }
            }

            SetAddReplaceButton(false, false);
            AddDetailsAsBox.ItemsSource = null;
        }

        private void ComboAddDetailsAs_DropDownClosed(object sender, EventArgs e)
        {
            int wi;

            if (int.TryParse(workItemID.Text, out wi))
            {
                if (AddDetailsAsBox.SelectedValue != null)
                {
                    string description = TfsManager.Instance.GetWorkItemDescription(wi, AddDetailsAsBox.SelectedValue.ToString());

                    if (description == string.Empty)
                        SetAddReplaceButton(true, false);
                    else
                        SetAddReplaceButton(false, true);

                    return;
                }
            }

            SetAddReplaceButton(false, false);
            //replaceButton.IsEnabled = (comboAddDetailsAs.Text != "Steps" && comboAddDetailsAs.Text != "Шаги");
        }

        private void filterBox_Scroll(object sender, System.Windows.Controls.Primitives.ScrollEventArgs e)
        {
        }
    }
}
