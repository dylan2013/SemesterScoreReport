using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using System.IO;
using Aspose.Words;
using System.Xml;
using SmartSchool.Common;

namespace SemesterScoreReport
{
    public partial class SemesterScoreReportConfig : System.Windows.Forms.Form
    {
        private byte[] _buffer = null;
        private string base64 = "";
        private bool _isUpload = false;

        public SemesterScoreReportConfig(bool useDefaultTemplate, byte[] buffer, int receiver, int address, string resitSign, string repeatSign)
        {
            InitializeComponent();

            if (useDefaultTemplate)
                radioButton1.Checked = true;
            else
                radioButton2.Checked = true;

            comboBoxEx1.SelectedIndex = receiver;
            comboBoxEx2.SelectedIndex = address;

            textBoxX1.Text = resitSign;
            textBoxX2.Text = repeatSign;

            _buffer = buffer;
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            #region 儲存 Preference

            XmlElement config = SmartSchool.Customization.Data.SystemInformation.Preference["_97入學_普綜_學年成績單"];

            if (config == null)
            {
                config = new XmlDocument().CreateElement("_97入學_普綜_學年成績單");
            }

            config.SetAttribute("UseDefault", (radioButton1.Checked ? "True" : "False"));

            XmlElement customize = config.OwnerDocument.CreateElement("CustomizeTemplate");
            XmlElement print = config.OwnerDocument.CreateElement("Print");

            if (_isUpload)
            {
                customize.InnerText = base64;
                config.ReplaceChild(customize, config.SelectSingleNode("CustomizeTemplate"));
            }

            print.SetAttribute("Name", comboBoxEx1.SelectedIndex.ToString());
            print.SetAttribute("Address", comboBoxEx2.SelectedIndex.ToString());
            print.SetAttribute("ResitSign", textBoxX1.Text);
            print.SetAttribute("RepeatSign", textBoxX2.Text);

            if (config.SelectSingleNode("Print") == null)
                config.AppendChild(print);
            else
                config.ReplaceChild(print, config.SelectSingleNode("Print"));

            SmartSchool.Customization.Data.SystemInformation.Preference["_97入學_普綜_學年成績單"] = config;

            #endregion

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
                radioButton2.Checked = false;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
                radioButton1.Checked = false;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "學期成績單.doc";
            sfd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(忠信學年成績單.Properties.Resources._97入學_普綜_學年成績單, 0, 忠信學年成績單.Properties.Resources._97入學_普綜_學年成績單.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "自訂學期成績單.doc";
            sfd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    if (_buffer != null && Aspose.Words.Document.DetectFileFormat(new MemoryStream(_buffer)) == Aspose.Words.LoadFormat.Doc)
                        fs.Write(_buffer, 0, _buffer.Length);
                    else
                    {
                        fs.Write(忠信學年成績單.Properties.Resources._97入學_普綜_學年成績單, 0, 忠信學年成績單.Properties.Resources._97入學_普綜_學年成績單.Length);
                    }
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "選擇自訂的學期成績單範本";
            ofd.Filter = "Word檔案 (*.doc)|*.doc";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (Document.DetectFileFormat(ofd.FileName) == LoadFormat.Doc)
                    {
                        FileStream fs = new FileStream(ofd.FileName, FileMode.Open);

                        byte[] tempBuffer = new byte[fs.Length];
                        fs.Read(tempBuffer, 0, tempBuffer.Length);
                        base64 = Convert.ToBase64String(tempBuffer);
                        _isUpload = true;
                        fs.Close();
                        MsgBox.Show("上傳成功。");
                    }
                    else
                        MsgBox.Show("上傳檔案格式不符");
                }
                catch
                {
                    MsgBox.Show("指定路徑無法存取。", "開啟檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }
    }
}