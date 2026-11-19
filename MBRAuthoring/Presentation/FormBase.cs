using MBRAuthoring.Business;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace MBRAuthoring.Presentation
{
    public partial class FormBase : Form
    {
        private int cboMaxWidth = 300;
        public static IFormManager _formManager;
        protected string currentDocName;
        protected List<TabPage> tabs = new List<TabPage>();
        protected TabControl tabControl;
        protected Button backButton;
        protected Button nextButton;
        protected Document currentDocument;
        protected virtual Panel progressBarPanel { get; }
        protected ToolTip toolTip;

        public static IFormManager formMangerInstance => _formManager;


        protected FormBase()
        {
            InitializeComponent();
            _formManager = new FormManager();
            // currentDocument = Globals.ThisDocument.ActiveDocument;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;
            toolTip = new ToolTip();
        }

        private void FormBase_Load(object sender, EventArgs e)
        {
            LimitCboBoxDropDownWidth(this);

        }

        private void LimitCboBoxDropDownWidth(Control control)
        {
            foreach (Control childControl in control.Controls)
            {

                if (childControl is ComboBox comboBox)
                {
                    comboBox.DropDown += (s, args) =>
                  {
                      if (comboBox.DropDownWidth > cboMaxWidth)
                      {
                          comboBox.DropDownWidth = cboMaxWidth;
                      }
                  };
                }

                // Recursively call LimitComboBoxDropdownWidth for child controls
                LimitCboBoxDropDownWidth(childControl);

            }
        }

        public void SetTabPageEnabledState(TabPage tabPage, bool isEnabled)
        {
            foreach (Control control in tabPage.Controls)
            {
                control.Enabled = isEnabled;
            }
        }

        public string PropertyBracketed(string propertyName)
        {
            return $"[{propertyName}]";
        }
        public void EnableSpecificControlsInTab(RadioButton fileOption, Dictionary<RadioButton, List<Control>> radioButtonMappings)
        {
            foreach (var radioButton in radioButtonMappings)
            {
                foreach (Control control in radioButton.Value)
                {
                    control.Visible = (radioButton.Key == fileOption);
                }
            }
        }

        public void cboBox_Format<T>(object sender, ListControlConvertEventArgs e, List<string> propertyNames, List<int> columnWidths)
        {
            if (e.ListItem is T item)
            {
                StringBuilder formattedValue = new StringBuilder();
                for (int i = 0; i < propertyNames.Count; i++)
                {
                    PropertyInfo prop = item.GetType().GetProperty(propertyNames[i]);
                    if (prop != null)
                    {
                        string value = prop.GetValue(item)?.ToString() ?? "";
                        //formattedValue.Append(value.PadRight(columnWidths[i]));
                        formattedValue.Append(value);
                    }
                    if (i < propertyNames.Count - 1)
                    {
                        formattedValue.Append("      ");
                    }
                }
                e.Value = formattedValue.ToString();
            }
        }


        protected void InitializeTabs(TabControl tabControl, Button backButton, Button nextButton, params TabPage[] tabPages)
        {
            tabs.AddRange(tabPages);
            this.tabControl = tabControl;
            this.backButton = backButton;
            this.nextButton = nextButton;
        }

        protected void ShowTab(int index)
        {
            foreach (var tab in tabs)
            {
                tabControl.TabPages.Remove(tab);
            }
            tabControl.TabPages.Add(tabs[index]);
            tabControl.SelectedIndex = 0;
            tabControl.Appearance = TabAppearance.Buttons;
            tabControl.SizeMode = TabSizeMode.Fixed;
            tabControl.ItemSize = new Size(0, 1);
            UpdateButtonStates();
        }



        private void UpdateButtonStates()
        {
            int currentIndex = GetCurrentTabIndex();
            nextButton.Enabled = currentIndex < tabs.Count - 1;
            backButton.Enabled = currentIndex > 0;

            nextButton.ForeColor = nextButton.Enabled ? Color.Black : Color.LightGray;
            backButton.ForeColor = backButton.Enabled ? Color.Black : Color.LightGray;
        }

        protected int GetCurrentTabIndex()
        {
            return tabs.IndexOf(tabControl.SelectedTab);
        }

        //protected void ShowNextTab()
        //{
        //    int currentIndex = GetCurrentTabIndex();
        //    int nextIndex = currentIndex + 1;
        //    if (nextIndex < tabs.Count)
        //    {
        //        ShowTab(nextIndex);
        //    }
        //}
        protected void ShowNextTab()
        {
            int currentIndex = GetCurrentTabIndex();
            int nextIndex = currentIndex + 1;
            if (nextIndex < tabs.Count)
            {
                ShowTab(nextIndex);
            }
            string currShpName = "shp" + nextIndex;
            string prevShpName = "shp" + currentIndex;
            Label currLabel = Controls.Find(currShpName, true).FirstOrDefault() as Label;
            Label prevLabel = Controls.Find(prevShpName, true).FirstOrDefault() as Label;
            if (currLabel != null) currLabel.BackColor = Color.Lime;
            if (prevLabel != null) prevLabel.BackColor = Color.Silver;
        }


        //protected void ShowPrevTab()
        //{
        //    int currentIndex = GetCurrentTabIndex();
        //    int prevIndex = currentIndex - 1;
        //    if (prevIndex >= 0)
        //    {
        //        ShowTab(prevIndex);
        //    }
        //}

        protected void ShowPrevTab()
        {
            int currentIndex = GetCurrentTabIndex();
            int prevIndex = currentIndex - 1;
            if (prevIndex >= 0)
            {
                ShowTab(prevIndex);
            }
            string currShpName = "shp" + prevIndex;
            string prevShpName = "shp" + currentIndex;
            Label currShp = Controls.Find(currShpName, true).FirstOrDefault() as Label;
            Label prevShp = Controls.Find(prevShpName, true).FirstOrDefault() as Label;
            if (currShp != null) currShp.BackColor = Color.Lime;
            if (prevShp != null) prevShp.BackColor = Color.Silver;
        }


        protected void ProgressLabel_Click(object sender, EventArgs e)
        {
            if (!(sender is Label clickedLabel)) return;

            if (!int.TryParse(clickedLabel.Tag?.ToString(), out int tabId)) return;

            ShowTab(tabId);

            foreach (Control ctrl in progressBarPanel.Controls)
            {
                if (ctrl is Label label)
                {
                    if (label.Tag != null || label.Tag.ToString() != "Exclude")
                    {
                        if (int.TryParse(label.Tag.ToString(), out int labelTagId))
                        {
                            if (label.Name.StartsWith("shp"))
                            {
                                label.BackColor = (labelTagId == tabId) ? Color.Lime : Color.Silver;
                            }
                        }
                    }
                }
            }
        }

        protected void SetToolTip(Control control, string tooltipText)
        {
            toolTip.BackColor = Color.FromArgb(255, 255, 225); // Light yellow
            toolTip.SetToolTip(control, tooltipText);
        }


    }
}
