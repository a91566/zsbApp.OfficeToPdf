using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace zsbApp.OfficeToPdf.Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dicOffice.Add("Office2016", "zsbApp.Office2010ToPdf.dll");
            dicOffice.Add("Office2013", "zsbApp.Office2010ToPdf.dll");
            dicOffice.Add("Office2010", "zsbApp.Office2010ToPdf.dll");
            dicOffice.Add("Office2007", "zsbApp.Office2007ToPdf.dll");
            dicOffice.Add("Office2003", "zsbApp.Office2003ToPdf.dll");
            dicOffice.Add("WPS2013", "zsbApp.WPS9ToPdf.dll");
            dicOffice.Add("WPS2016", "zsbApp.WPS10ToPdf.dll");
            dicOffice.Add("默认", "zsbApp.CrackedAspose17ToPdf.dll");

            cmdOffice.Items.AddRange(dicOffice.Keys.ToArray());
            cmdOffice.SelectedValueChanged += new EventHandler(cmdOffice_SelectedValueChanged);
        }

        void cmdOffice_SelectedValueChanged(object sender, EventArgs e)
        {
            if (dicOffice.ContainsKey(cmdOffice.Text))
            {
                txbDllName.Text = dicOffice[cmdOffice.Text];
            }
        }

        private zsbApp.OfficeToPdf.IOfficeToPdf iOfficeToPdf;

        Dictionary<string, string> dicOffice = new Dictionary<string, string>();

        private void btnSelectInterface_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "可执行文件|*.dll;*.exe;";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                this.txbInterface.Text = openFileDialog.FileName;
                this.createIOfficeToPdf(openFileDialog.FileName);
                if (this.iOfficeToPdf != null)
                {
                    this.lblInterfaceTag.Text = this.iOfficeToPdf.Tag;
                }
                else
                {
                    MessageBox.Show("没有找到对应的接口");
                }
            }
        }

        private void btnSelectOfficeFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "office 文档 |*.doc;*.docx;*.xls;*.xlsx;";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = openFileDialog.FileName;
                this.txbOfficeFile.Text = fileName;
                if (fileName.EndsWith(".docx"))
                {
                    this.txbPdf.Text = fileName.Replace(".docx", ".pdf");
                }
                else if (fileName.EndsWith(".doc"))
                {
                    this.txbPdf.Text = fileName.Replace(".doc", ".pdf");
                }
                else if (fileName.EndsWith(".xls"))
                {
                    this.txbPdf.Text = fileName.Replace(".xls", ".pdf");
                }
                else if (fileName.EndsWith(".xlsx"))
                {
                    this.txbPdf.Text = fileName.Replace(".xlsx", ".pdf");
                }
            }
        }

        private void createIOfficeToPdf(string file)
        {
            var dllFile = Path.Combine(Application.StartupPath, file);
            var asm = System.Reflection.Assembly.LoadFile(dllFile);
            foreach (var type in asm.GetTypes())
            {
                if (type.GetInterfaces().Contains(typeof(zsbApp.OfficeToPdf.IOfficeToPdf)))
                {
                    this.iOfficeToPdf = Activator.CreateInstance(type) as zsbApp.OfficeToPdf.IOfficeToPdf;
                    if (this.iOfficeToPdf != null)
                        break;
                }
            }
        }

        private void btnToPdf_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(this.txbOfficeFile.Text) || string.IsNullOrEmpty(this.txbPdf.Text))
            {
                MessageBox.Show("文件信息未完整");
                return;
            }
            if (this.iOfficeToPdf == null)
            {
                MessageBox.Show("没有正确的接口");
                return;
            }

            var result = this.iOfficeToPdf.Convert(this.txbOfficeFile.Text, this.txbPdf.Text);
            if (string.IsNullOrEmpty(result))
            {
                System.Diagnostics.Process.Start(this.txbPdf.Text);
                return;
            }
            else
            {
                MessageBox.Show(result); 
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string> result = new List<string>();
            Type type1 = Type.GetTypeFromProgID("ET.Application");//V8版本类型
            if (type1 == null)
            {
                result.Add("没有安装V8版本");
            }
            else
            {
                result.Add("有安装V8版本");
            }


            Type type2 = Type.GetTypeFromProgID("Ket.Application");//V9版本类型
            if (type2 == null)
            {
                result.Add("没有安装了 WPS2013");
            }
            else
            {
                result.Add("安装了 WPS2013");
            }


            Type type3 = Type.GetTypeFromProgID("Kwps.Application");//V10版本类型
            if (type3 == null)
            {
                result.Add("没有安装 WPS2016");
            }
            else
            {
                result.Add("安装了 WPS2016");
            }

            Type type4 = Type.GetTypeFromProgID("EXCEL.Application");//MS EXCEL类型
            if (type4 == null)
            {
                result.Add("没有安装了 Office");
            }
            else
            {
                result.Add("安装了 Office");
            }

            this.textBox1.Text = string.Join(Environment.NewLine, result);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            RegeditTool RegeditTool = new RegeditTool();
            var office = RegeditTool.ExistsRegedit();
  
            if (!string.IsNullOrEmpty(office) && dicOffice.ContainsKey(office))
            {
                txbOfficeFile.Text = dicOffice[office];
                MessageBox.Show("检测到使用的是" + office);
            }
            else
            {
                MessageBox.Show("没有检测到安装了OFFICE或者WPS软件,将使用默认方式");

                txbOfficeFile.Text = dicOffice["默认"];
            }
            this.createIOfficeToPdf(txbOfficeFile.Text);
            btnSelectOfficeFile_Click(null, null);
            btnToPdf_Click(null, null);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.createIOfficeToPdf(txbDllName.Text);
            btnSelectOfficeFile_Click(null, null);
            btnToPdf_Click(null, null);
        }


    }
}
