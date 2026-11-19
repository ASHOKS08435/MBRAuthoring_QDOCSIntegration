using System;
using System.Windows.Forms;
using System.Diagnostics;
using MBRAuthoring.Business;


namespace MBRAuthoring.Presentation
{
    public class InputBox : FormBase
    {

        private TextBox textBox;
        private Button btnOk;
        private Button btnCancel;
        string mstrProtectionPassword;
        string MyCompName = "frmProtection";
        private bool isClosingFromButton = false;

        public string Input { get; private set; }
        public InputBox(string prompt, string protectionPassword, string docName)
        {
            //this.FormBorderStyle = FormBorderStyle.FixedSingle;
            //this.MaximizeBox = false;
            //this.MinimizeBox = false;
            //this.ShowIcon = false;
            //InitializeComponent();
            mstrProtectionPassword = protectionPassword;

            this.Width = 360;
            this.Height = 170;

            Label lblPrompt = new Label() { Left = 30, Top = 20, Text = prompt, AutoSize = true };
            lblPrompt.Text = "Password for " + docName;


            textBox = new TextBox { Left = 30, Top = 90, Width = 300 };
            this.Text = MyCompName;
            textBox.Text = mstrProtectionPassword;
            textBox.Enabled = true;

            btnOk = new Button { Text = "OK", Left = 270, Top = 20, Width = 60, DialogResult = DialogResult.OK };
            btnCancel = new Button { Text = "Cancel", Left = 270, Top = 50, Width = 60, DialogResult = DialogResult.Cancel };

            btnOk.Click += new EventHandler(btnOk_Click);
            btnCancel.Click += new EventHandler(btnCancel_Click);
            this.FormClosing += InputForm_FormClosing;

            this.Controls.Add(lblPrompt);
            this.Controls.Add(textBox);
            this.Controls.Add(btnOk);
            this.Controls.Add(btnCancel);
            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
            this.StartPosition = FormStartPosition.CenterParent;

        }

        public void btnOk_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            isClosingFromButton = true;

            if (textBox.Text != mstrProtectionPassword)
            {
                MessageBox.Show("Note: Password display is for your information only," + Environment.NewLine + "password change is ignored.", MyCompName, MessageBoxButtons.OK);
            }
            this.Close();
            CommonService.CleanUpBackgroundWordProcesses();

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            isClosingFromButton = true;
            MessageBox.Show("Note: Password display is for your information only," + Environment.NewLine + "password change is ignored.", MyCompName, MessageBoxButtons.OK);
            this.Close();
            CommonService.CleanUpBackgroundWordProcesses();

        }


        public static string ShowDialog(string prompt, string title, string protectionPassword = "", string docName = "")
        {
            using (InputBox box = new InputBox(prompt, protectionPassword, docName))
            {
                box.Text = title;
                return box.ShowDialog() == DialogResult.OK ? box.Input : null;
            }
        }

        private void InputForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Visible = false;

            if (isClosingFromButton)
            {
                return;
            }

            if (e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult result = MessageBox.Show("Note: Password display is for your information only," + Environment.NewLine + "password change is ignored.", MyCompName, MessageBoxButtons.OK);

                if (result == DialogResult.Cancel)
                {
                    e.Cancel = true; // Prevent the form from closing
                }
            }

        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // InputBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Name = "InputBox";
            this.ResumeLayout(false);

        }
    }
}
