using System;

namespace MBRAuthoring.Presentation
{
    public partial class FormMBRHelp : FormBase
    {
        public FormMBRHelp()
        {
            InitializeComponent();

        }

        private void FormMBRHelp_Load(object sender, EventArgs e)
        {

            InitializeForm();
        }

        private void InitializeForm()
        {

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Close();
            
        }
    }
}
