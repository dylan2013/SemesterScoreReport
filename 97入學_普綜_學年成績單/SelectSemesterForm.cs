using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using SmartSchool.Customization.Data;

namespace SemesterScoreReport
{
    public partial class SelectSemesterForm : System.Windows.Forms.Form
    {
        protected int _schoolyear = 0;
        protected int _semester = 0;

        public int SchoolYear
        {
            get { return _schoolyear; }
        }
        public int Semester
        {
            get { return _semester; }
        }

        public SelectSemesterForm(string titleName)
            : this()
        {
            this.Text = titleName;
        }

        public SelectSemesterForm()
        {
            InitializeComponent();

            //if ( !DesignMode ) return;
            try
            {
                numericUpDown1.Value = decimal.Parse(SmartSchool.Customization.Data.SystemInformation.SchoolYear.ToString());
                numericUpDown2.Value = decimal.Parse(SmartSchool.Customization.Data.SystemInformation.Semester.ToString());
            }
            catch { }
        }

        protected void buttonX1_Click(object sender, EventArgs e)
        {
            _schoolyear = int.Parse(numericUpDown1.Value.ToString());
            _semester = int.Parse(numericUpDown2.Value.ToString());

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        protected void buttonX2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}