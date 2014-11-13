using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using WindowsFormsControlLibrary1;

using OpcRcw.Da;
using Microsoft.VisualBasic.FileIO;

namespace PCDCS_Testing_Tool
{
    public partial class Form1 : Form
    {
        DxpSimpleAPI.DxpSimpleClass opc = new DxpSimpleAPI.DxpSimpleClass();
        List<string> registers = new List<string>() { };
        List<string> tags = new List<string>() { };
        List<TextBox> tag_no = new List<TextBox>();
        List<TextBox> reg_no = new List<TextBox>();
        List<TextBox> valueReg = new List<TextBox>();
        string[] sItemIDArray = new string[5];
        int a = -1;
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void btnListRefresh_Click(object sender, EventArgs e)
        {
            cmbServerList.Items.Clear();
            string[] ServerNameArray;
            opc.EnumServerList(txtNode.Text, out ServerNameArray);

            for (int a = 0; a < ServerNameArray.Count<string>(); a++)
            {
                cmbServerList.Items.Add(ServerNameArray[a]);
            }
            cmbServerList.SelectedIndex = 0;
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            if (opc.Connect(txtNode.Text, cmbServerList.Text))
            {
                btnListRefresh.Enabled = false;
                btnDisconnect.Enabled = true;
                btnConnect.Enabled = false;

            }
        }

        object[] oValue = new object[5];
               
        private object WriteValCopy(string sText, int n)
        {
            if (oValue[n] is Array)
            {
                string[] sBufPut = sText.Split(',');
                Array ary = (Array)oValue[n];
                object[] oAry = new object[ary.Length];
                for (int j = 0; j < ary.Length; j++)
                {
                    if (j >= sBufPut.Length)
                        break;
                    try
                    {
                        if (oValue[n] is UInt16)
                        {
                            oAry[j] = UInt16.Parse(sBufPut[j]);
                        }
                        else if (oValue[n] is UInt32)
                        {
                            oAry[j] = UInt32.Parse(sBufPut[j]);
                        }
                        else
                        {
                            oAry[j] = (object)sBufPut[j];
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Error");
                    }
                }
                return oAry;
            }
            else
            {
                return sText;
            }
        }

        private void FileReadBtn_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                List<string[]> list = new List<string[]>();
                using (TextFieldParser parser = new TextFieldParser(openFileDialog1.FileName, Encoding.GetEncoding(932)))
                {
                    parser.Delimiters = new string[] { "," };
                    bool st = false;

                    string checkPointString = "";
                    if (Properties.Settings.Default.EnglishMode == true)
                    {
                        checkPointString = "[Dictionary for Alarm]";
                    }
                    else
                    {
                        checkPointString = "[アラーム用辞書]";
                    }

                    while (true)
                    {
                        string[] parts = parser.ReadFields();
                        if (parts == null)
                        {
                            break;
                        }
                        if (st)
                        {
                            if (parts[0] != "[END]")
                            {
                                list.Add(parts);
                            }
                            else
                            {
                                st = false;
                            }
                        }

                        if (parts[0] == checkPointString)
                        {
                            st = true;
                        }
                    }


                    maintable.Controls.Clear();

                    if (list.Count > 0)
                    {
                        registers.Clear();
                        tags.Clear();
                        tag_no.Clear();
                        reg_no.Clear();
                        valueReg.Clear();
                        maintable.Visible = false;
                        maintable.RowCount = list.Count;
                        maintable.Height = 29 * list.Count;
                        for (int i = 0; i < list.Count; i++)                   
                        {
                            TextBox txtNo = new TextBox();
                            txtNo.Text = i.ToString();
                            txtNo.ReadOnly = true;
                            maintable.Controls.Add(txtNo, 0, i);

                            TextBox txtTagNo = new TextBox();
                            txtTagNo.Text = list[i][1];
                            txtTagNo.ReadOnly = true;
                            maintable.Controls.Add(txtTagNo, 1, i);
                            tag_no.Add(txtTagNo);

                            TextBox txtRegNo = new TextBox();
                            txtRegNo.Text = list[i][6];
                            txtRegNo.ReadOnly = true;
                            maintable.Controls.Add(txtRegNo, 2, i);
                            reg_no.Add(txtRegNo);

                            TextBox txtStatus = new TextBox();
                            maintable.Controls.Add(txtStatus, 3, i);
                            valueReg.Add(txtStatus);

                            Button btnOn = new Button();
                            btnOn.Text = "On";
                            btnOn.Tag = i.ToString();
                            btnOn.Click += btnOn_Click;
                            maintable.Controls.Add(btnOn, 4, i);


                            Button btnOff = new Button();
                            btnOff.Text = "Off";
                            btnOff.Tag = i.ToString();
                            btnOff.Click += btnOff_Click;
                            maintable.Controls.Add(btnOff, 5, i);

                            registers.Add(list[i][6]);
                            tags.Add(list[i][1]);
                        }
                        maintable.Visible = true;
                        button1.Enabled = true;
                    }
                    else {
                        Label message = new Label();
                        maintable.Visible = false;
                        message.Text = "There are no lists inside the file.";
                        message.Location = new Point(0, 30);
                        message.Width = 200;                        
                        panel1.Controls.Add(message);
                    }
                }
            }
        }

        void btnOff_Click(object sender, EventArgs e)
        {
            OpcOnOff(0, Convert.ToInt32((sender as Button).Tag.ToString()));
        }

        void btnOn_Click(object sender, EventArgs e)
        {
            OpcOnOff(1, Convert.ToInt32((sender as Button).Tag.ToString()));
        }

        private void btnDisconnect_Click(object sender, EventArgs e)
        {
            if (opc.Disconnect())
            {
                btnConnect.Enabled = true;
                btnListRefresh.Enabled = true;
                btnDisconnect.Enabled = false;

                //txtTag1.ReadOnly = false;
                //txtTag2.ReadOnly = false;
                //txtTag3.ReadOnly = false;
                //txtTag4.ReadOnly = false;
                //txtTag5.ReadOnly = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (a > -1)
            {
                reg_no[a].BackColor = SystemColors.Control;
                tag_no[a].BackColor = SystemColors.Control;
            }
            if (registers.Contains(txtReg.Text))
            {
                a = registers.IndexOf(txtReg.Text);
                reg_no[a].BackColor = Color.Red;
                panel1.VerticalScroll.Value = a * 29;
            }
            else if (tags.Contains(txtReg.Text))
            {
                a = tags.IndexOf(txtReg.Text);
                tag_no[a].BackColor = Color.Red;
                panel1.VerticalScroll.Value = a * 29;
            }
        }

        private void OpcOnOff(int value, int tag)
        {
            try
            {
                string[] target = new string[] { reg_no[tag].Text };
                object[] val = new object[] { value };
                int[] nErrorArray;

                data1.ColumnCount = 5;
                data1.Columns[0].Name = "Date Time";
                data1.Columns[1].Name = "Tag No.";
                data1.Columns[2].Name = "Register No";
                data1.Columns[3].Name = "Status";
                data1.Columns[4].Name = "Success/Error";

                //StreamWriter sw = new StreamWriter(file, true);
                if (opc.Write(target, val, out nErrorArray))
                {
                    //sw.WriteLine(DateTime.Now.ToString() + "," + tag_no[tag].Text + "," + reg_no[tag].Text + "," + "Write,Success");
                    string[] row = new string[] { DateTime.Now.ToString(), tag_no[tag].Text, reg_no[tag].Text, valueReg[tag].Text, "Write Success" };
                    data1.Rows.Add(row);
                }
                else
                {
                    valueReg[tag].Text = "Write Error";
                    //sw.WriteLine(DateTime.Now.ToString() + "," + tag_no[tag].Text + "," + reg_no[tag].Text + "," + "Write,Error");
                    string[] row = new string[] { DateTime.Now.ToString(), tag_no[tag].Text, reg_no[tag].Text, valueReg[tag].Text, "Write Error" };
                    data1.Rows.Add(row);
                }
                short[] wQualityArray;
                OpcRcw.Da.FILETIME[] fTimeArray;

                if (opc.Read(target, out val, out wQualityArray, out fTimeArray, out nErrorArray) == true)
                {
                    valueReg[tag].Text = val[0].ToString();
                    //sw.WriteLine(DateTime.Now.ToString() + "," + tag_no[tag].Text + "," + reg_no[tag].Text + "," + "Read,Success");
                    string[] row = new string[] { DateTime.Now.ToString(), tag_no[tag].Text, reg_no[tag].Text, valueReg[tag].Text, "Read Success" };
                    data1.Rows.Add(row);
                }
                else
                {

                    //sw.WriteLine(DateTime.Now.ToString() + "," + tag_no[tag].Text + "," + reg_no[tag].Text + "," + "Read,Error");
                    string[] row = new string[] { DateTime.Now.ToString(), tag_no[tag].Text, reg_no[tag].Text, valueReg[tag].Text, "Read Error" };
                    data1.Rows.Add(row);
                }
                //sw.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

  
    }
}
