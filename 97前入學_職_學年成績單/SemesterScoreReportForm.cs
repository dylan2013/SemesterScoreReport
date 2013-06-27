using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using SmartSchool.Common;

namespace SemesterScoreReport
{
    public partial class SemesterScoreReportFormNew : SelectSemesterForm
    {
        private Dictionary<string, List<string>> _userType = new Dictionary<string, List<string>>();
        private MemoryStream _template = null;
        private bool _useDefaultTemplate = false;
        private int _receiver = 0;
        private int _address = 0;
        private byte[] _buffer = null;
        private string _resitSign = "*";
        private string _repeatSign = "#";

        public Dictionary<string, List<string>> UserDefinedType
        {
            get { return _userType; }
        }

        public MemoryStream Template
        {
            get
            {
                if (_useDefaultTemplate)
                    return new MemoryStream(忠信學年成績單.Properties.Resources._97前入學_職_學年成績單);
                else if (_template != null)
                    return _template;
                else
                    return new MemoryStream(忠信學年成績單.Properties.Resources._97前入學_職_學年成績單);
            }
        }

        public int Receiver { get { return _receiver; } }
        public int Address { get { return _address; } }
        public string ResitSign { get { return _resitSign; } }
        public string RepeatSign { get { return _repeatSign; } }

        public SemesterScoreReportFormNew()
        {
            InitializeComponent();
            LoadPreference();
        }

        private void LoadPreference()
        {
            XmlElement config = SmartSchool.Customization.Data.SystemInformation.Preference["_97前入學_職_學年成績單"];
            if (config != null)
            {
                //使用者設定的假別
                _userType.Clear();

                foreach (XmlElement type in config.SelectNodes("Type"))
                {
                    string typeName = type.GetAttribute("Text");

                    if (!_userType.ContainsKey(typeName))
                        _userType.Add(typeName, new List<string>());

                    foreach (XmlElement absence in type.SelectNodes("Absence"))
                    {
                        string absenceName = absence.GetAttribute("Text");

                        if (!_userType[typeName].Contains(absenceName))
                            _userType[typeName].Add(absenceName);
                    }
                }

                //範本
                if (config.HasAttribute("UseDefault"))
                    _useDefaultTemplate = bool.Parse(config.GetAttribute("UseDefault"));
                else
                {
                    config.SetAttribute("UseDefault", "True");
                    SmartSchool.Customization.Data.SystemInformation.Preference["_97前入學_職_學年成績單"] = config;
                }

                XmlElement customize = (XmlElement)config.SelectSingleNode("CustomizeTemplate");

                if (customize != null)
                {
                    if (!string.IsNullOrEmpty(customize.InnerText))
                    {
                        string templateBase64 = customize.InnerText;
                        _buffer = Convert.FromBase64String(templateBase64);
                        _template = new MemoryStream(_buffer);
                    }
                }
                else
                {
                    XmlElement newCustomize = config.OwnerDocument.CreateElement("CustomizeTemplate");
                    config.AppendChild(newCustomize);
                    SmartSchool.Customization.Data.SystemInformation.Preference["_97前入學_職_學年成績單"] = config;
                }

                //列印資訊
                XmlElement print = (XmlElement)config.SelectSingleNode("Print");
                if (print != null)
                {
                    _receiver = int.Parse(print.GetAttribute("Name"));
                    _address = int.Parse(print.GetAttribute("Address"));
                    _resitSign = print.GetAttribute("ResitSign");
                    _repeatSign = print.GetAttribute("RepeatSign");
                }
                else
                {
                    XmlElement newPrint = config.OwnerDocument.CreateElement("Print");
                    newPrint.SetAttribute("Name", "0");
                    newPrint.SetAttribute("Address", "0");
                    newPrint.SetAttribute("ResitSign", "*");
                    newPrint.SetAttribute("RepeatSign", "#");
                    _receiver = 0;
                    _address = 0;
                    _resitSign = "*";
                    _repeatSign = "#";
                    config.AppendChild(newPrint);
                    SmartSchool.Customization.Data.SystemInformation.Preference["_97前入學_職_學年成績單"] = config;
                }
            }
            else
            {
                #region 產生空白設定檔
                config = new XmlDocument().CreateElement("_97前入學_職_學年成績單");
                SmartSchool.Customization.Data.SystemInformation.Preference["_97前入學_職_學年成績單"] = config;
                #endregion
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SemesterScoreReportConfig configForm = new SemesterScoreReportConfig(_useDefaultTemplate, _buffer, _receiver, _address, _resitSign, _repeatSign);
            if (configForm.ShowDialog() == DialogResult.OK)
            {
                LoadPreference();
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SelectTypeForm form = new SelectTypeForm("_97前入學_職_學年成績單");
            if (form.ShowDialog() == DialogResult.OK)
            {
                LoadPreference();
            }
        }

        private void buttonX1_Click_1(object sender, EventArgs e)
        {

        }
    }
}