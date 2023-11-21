using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using GeneralUtility;
using MySql.Data.MySqlClient;
using System.Web;
using System.IO;

namespace MaterialＭanager
{
    public partial class MForm : Form
    {
        private MySqlConnection _connection;
        private string _dbHost = "127.0.0.1";
        private string _dbPort = "3306";
        private string _dbUserName = "root";
        private string _dbPassword = "";
        private string _dbName = "invent";

        private XXLog _clsLog = new XXLog(Application.StartupPath + "\\Log");

        private Thread _thd_inOK = null;
        private Thread _thd_outOK = null;

        int countvalue_in;
        int countvalue_out;
        string strogebox_in;
        string strogebox_out;

        public MForm()
        {
            InitializeComponent();
        }

        private void MForm_Load(object sender, EventArgs e)
        {
            string connStr = string.Format("server={0}; port={1}; uid={2}; pwd={3}; database={4}; charset=utf8;", _dbHost, _dbPort, _dbUserName, _dbPassword, _dbName);
            
            _connection = new MySqlConnection(connStr);
            _connection.Open();

            usertextBox.Focus();
            clean(null, null);
            _clsLog.Info("Application Started...");
        }

        private void MForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            _clsLog.Info("Applicaiton Closed...");
            _connection.Close();
            Thread.Sleep(500);
            _clsLog.Close(); 
        }

        private void clean(object sender, KeyEventArgs e)
        {
            ShowProductlabel.Text = "";
            ShowPalletlabel.Text = "";
            ShowPalletLocationlabel.Text = "";
            ShowPanelCountlabel.Text = "";
            ShowTrayBoxlabel.Text = "";
            ShowCommentlabel.Text = "";
            ShowProductOlabel.Text = "";
            ShowPalletOlabel.Text = "";
            ShowLocationlabel.Text = "";
            ShowPanelCountOlabel.Text = "";
            ShowTrayBoxOlabel.Text = "";
            ShowCommentOlabel.Text = "";
            listBox2.Items.Clear();
            listBox3.Items.Clear();
            totallabel_in.Text = "";
            totallabel_out.Text = "";
            searchlabel.Text = "";
            CountResulelabel.Text = "";

            Plabel11.Text = Plabel12.Text = Plabel13.Text = Plabel14.Text = "";
            Plabel21.Text = Plabel22.Text = Plabel23.Text = Plabel24.Text = "";
            Plabel31.Text = Plabel32.Text = Plabel33.Text = Plabel34.Text = "";
            Plabel41.Text = Plabel42.Text = Plabel43.Text = Plabel44.Text = "";
            Plabel51.Text = Plabel52.Text = Plabel53.Text = Plabel54.Text = "";
            Plabel61.Text = Plabel62.Text = Plabel63.Text = Plabel64.Text = "";
            Plabel71.Text = Plabel72.Text = Plabel73.Text = Plabel74.Text = "";
            Plabel81.Text = Plabel82.Text = Plabel83.Text = Plabel84.Text = "";
            Plabel91.Text = Plabel92.Text = Plabel93.Text = Plabel94.Text = "";
            Plabel101.Text = Plabel102.Text = Plabel103.Text = Plabel104.Text = "";
            Plabel111.Text = Plabel112.Text = Plabel113.Text = Plabel114.Text = "";
            Plabel121.Text = Plabel122.Text = Plabel123.Text = Plabel124.Text = "";
            Plabel131.Text = Plabel132.Text = Plabel133.Text = Plabel134.Text = "";
            Plabel141.Text = Plabel142.Text = Plabel143.Text = Plabel144.Text = "";
            Plabel151.Text = Plabel152.Text = Plabel153.Text = Plabel154.Text = "";
            Plabel161.Text = Plabel162.Text = Plabel163.Text = Plabel164.Text = "";
        }
        


        private void tb_userid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) 
            {
                btn_signinout_Click(null, null); 
            }
        }

        private void btn_signinout_Click(object sender, EventArgs e)
        {
            if (usertextBox.Enabled == false)
            {
                _clsLog.Info(usertextBox.Text + ",  Sign out application."); 
                usertextBox.Text = "";
                usertextBox.Enabled = true;
                usertextBox.Focus();
                userbutton.Text = "啟用";
                tabcontrol.Enabled = false;
            }
            else 
            {
                if (usertextBox.Text.Length < 8 || usertextBox.Text.Length > 8)
                {
                    usertextBox.Text = "";
                    return;
                }

                string strID = usertextBox.Text.Substring(usertextBox.Text.Length - 8);
                usertextBox.Text = strID;
                countvalue_in = countvalue_out = 0;
                usertextBox.Enabled = false;
                userbutton.Text = "登出"; 
                tabcontrol.Enabled = true;
                //trayboxIDtextBox.Focus(); 
                productIDtextBox.Focus();
                _clsLog.Info(usertextBox.Text + ", Sign in application.");
                clean(null, null);
            }
        }

        private void tb_boxnumber_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                string strboxs = keyintextBox.Text;
                string stroper = usertextBox.Text;
                string strproduct = productIDtextBox.Text;
                string strpallet = palletIDtextBox.Text;
                string strlocation = palletlocationtextBox.Text;
                string strcount = counttextBox.Text;
                string strcomment = commenttextBox.Text;
                btn_in_Click(null, null);
                trayboxIDtextBox.Text = "";
                trayboxIDtextBox.Focus();
                _clsLog.Info(String.Format("{0}, Input Error Tray/Box count... , {1}, {2}, {3}, {4}, {5} ,{6}", stroper, strboxs, strproduct, strpallet, strlocation, strcount, strcomment));
                //_clsLog.Info(usertextBox.Text + " , Input Error Tray/Box count...");
            }
        }

        private void tb_productid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                palletIDtextBox.Focus();
            }
        }

        private void tb_palletnumber_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                palletlocationtextBox.Focus();
            }
        }

        private void tb_location_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                counttextBox.Focus();
            }
        }

        private void tb_comment_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                trayboxIDtextBox.Focus();
            }
        }

        private void BoxInOK()
        {
            setLabelVisible(label12, true);
            clearTabpageIn();
            Thread.Sleep(3000);
            setLabelVisible(label12, false);
        }

        private void clearTabpageIn()
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                    trayboxIDtextBox.Text = "";
                }));
            }
            else
            {
                trayboxIDtextBox.Text = "";
                productIDtextBox.Text = "";
                palletlocationtextBox.Text = "";
                palletIDtextBox.Text = "";
                counttextBox.Text = "";
                commenttextBox.Text = "";
                trayboxIDtextBox.Focus();
            }
        }

        private void btn_in_Click(object sender, EventArgs e)
        {
            string strbox = trayboxIDtextBox.Text;
            string strproduct = productIDtextBox.Text;
            string strpallet = palletIDtextBox.Text;
            string strlocation = palletlocationtextBox.Text;
            string strcount = counttextBox.Text;
            string strcomment = commenttextBox.Text.Trim(new char[] { ' ', '\r', '\n' });
            string stroper = usertextBox.Text;

            if (strbox == "" || strproduct == "" || strpallet == "" )
      
            {
                MessageBox.Show("資料不完整，清重新輸入。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            if (!_connection.Ping()) _connection.Open();

            if (_connection.Ping())
            {
                MySqlCommand cmd = _connection.CreateCommand();

                int icount = GetSQLCount(_connection, "boxid", strbox);

                if (icount > 0)
                {
                    cmd.CommandText = string.Format("UPDATE items SET productid = '{0}', palletid = '{1}', date_in = '{2}', date_out = NULL, last_oper = '{3}', _Location = '{4}', `comment` = '{5}',_count = '{6}' WHERE boxid = '{7}';",
                            strproduct, strpallet, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), stroper, strlocation, strcomment, strcount, strbox);
                    cmd.CommandText += string.Format("UPDATE items SET _location = '{0}' WHERE palletid = '{1}';", strlocation, strpallet);
                }
                else
                {
                    cmd.CommandText = string.Format("INSERT INTO items ( boxid, productid, palletid, date_in, last_oper, _location,_count,`comment`) VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}');",
                        strbox, strproduct, strpallet, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), stroper, strlocation, strcount, strcomment);
                    cmd.CommandText += string.Format("UPDATE items SET _location = '{0}' WHERE palletid = '{1}';", strlocation, strpallet);
                }

                try
                {
                    cmd.ExecuteNonQuery();
                    _clsLog.Info(String.Format("{0}, 入料 , {1}, {2}, {3}, {4}, {5} ,{6}", stroper, strbox, strproduct, strpallet, strlocation ,strcount ,strcomment));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                MessageBox.Show("資料庫連線異常，請重新嘗試。", "資料庫異常", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            ShowProductlabel.Text = strproduct;
            ShowPalletlabel.Text = strpallet;
            ShowPalletLocationlabel.Text = strlocation;
            ShowPanelCountlabel.Text = strcount;
            ShowTrayBoxlabel.Text = strbox;
            ShowCommentlabel.Text = strcomment;
            if(this.listBox3.Items.Count > 0)
            {
                if (!listBox3.Items.Contains(strbox))
                {
                    listBox3.Items.Add(strbox);
                }
            }
            else
            {
                listBox3.Items.Add(strbox);
            }

            _thd_inOK = new Thread(BoxInOK);
            _thd_inOK.IsBackground = true;
            _thd_inOK.Start();

            totallabel_in.Text = listBox3.Items.Count.ToString();
            strogebox_in = trayboxIDtextBox.Text;
        }

        private void btn_clearin_Click(object sender, EventArgs e)
        {
            clearTabpageIn();
        }

       private void cmb_searchcolumn_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == -1) return;
            searchtextBox.Focus(); 
        }

        private void tb_searchvalue_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_search_Click(null, null);
            }
        }

        private void btn_search_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            DataTable dt_buffer = new DataTable();
            
            if (comboBox2.SelectedIndex == -1 || searchtextBox.Text == "")
            {
                if (!_connection.Ping()) _connection.Open();
                if (_connection.Ping())
                {
                    string strCommand = "select * from items where date_out is null";
                    MySqlDataAdapter adapter = new MySqlDataAdapter(strCommand, _connection);
                    adapter.Fill(dt_buffer);
                }
            }
            else
            {
                string strkey = "";
                switch (comboBox2.Text.ToUpper())
                {
                    case "PRODUCT ID":
                        strkey = "productid";
                        break;
                    case "BOX NUMBER":
                        strkey = "boxid";
                        break;
                    case "PALLET NUMBER":
                        strkey = "palletid";
                        break;
                }

                string strval = searchtextBox.Text;

                if (!_connection.Ping()) _connection.Open();
                if (_connection.Ping())
                {
                    string strCommand = string.Format("select * from items where {0} = '{1}' and date_out is null", strkey, strval);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(strCommand, _connection);
                    adapter.Fill(dt_buffer);
                }
            }
            LoadDataTableToDGV(dt_buffer);
        }

        private void LoadDataTableToDGV(DataTable dt) 
        {
            if (dt== null || dt.Rows.Count == 0) return;
            foreach (DataRow dr in dt.Rows) 
            {
                int irowindex = dataGridView2.Rows.Add(dr);
                foreach (DataGridViewColumn col in dataGridView2.Columns) 
                {
                    switch (col.Name) 
                    {
                        case "col_checked":
                            dataGridView2.Rows[irowindex].Cells["col_checked"].Value = false;
                            break; 
                        case "col_boxnumber":
                            dataGridView2.Rows[irowindex].Cells["col_boxnumber"].Value = dr["boxid"].ToString(); 
                            break;
                        case "col_palletnumber":
                            dataGridView2.Rows[irowindex].Cells["col_palletnumber"].Value = dr["palletid"].ToString();
                            break;
                        case "col_productid":
                            dataGridView2.Rows[irowindex].Cells["col_productid"].Value = dr["productid"].ToString();
                            break;
                        case "col_location":
                            dataGridView2.Rows[irowindex].Cells["col_location"].Value = dr["_location"].ToString();
                            break;
                        case "col_date_in":
                            dataGridView2.Rows[irowindex].Cells["col_date_in"].Value = dr["date_in"].ToString();
                            break;
                        case "col_comment":
                            dataGridView2.Rows[irowindex].Cells["col_comment"].Value = dr["comment"].ToString();
                            break; 
                    }
                }
            }
        }

        private void LoadDataTableToDGVChange(DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0) return;
            foreach (DataRow dr in dt.Rows)
            {
                int irowindex = dataGridView5.Rows.Add(dr);
                foreach (DataGridViewColumn col in dataGridView5.Columns)
                {
                    switch (col.Name)
                    {
                        case "change_col_checked":
                            dataGridView5.Rows[irowindex].Cells["change_col_checked"].Value = false;
                            break;
                        case "change_col_boxnumber":
                            dataGridView5.Rows[irowindex].Cells["change_col_boxnumber"].Value = dr["boxid"].ToString();
                            break;
                        case "change_col_palletnumber":
                            dataGridView5.Rows[irowindex].Cells["change_col_palletnumber"].Value = dr["palletid"].ToString();
                            break;
                        case "change_col_productid":
                            dataGridView5.Rows[irowindex].Cells["change_col_productid"].Value = dr["productid"].ToString();
                            break;
                        case "change_col_location":
                            dataGridView5.Rows[irowindex].Cells["change_col_location"].Value = dr["_location"].ToString();
                            break;
                        case "change_col_date_in":
                            dataGridView5.Rows[irowindex].Cells["change_col_date_in"].Value = dr["date_in"].ToString();
                            break;
                        case "change_col_comment":
                            dataGridView5.Rows[irowindex].Cells["change_col_comment"].Value = dr["comment"].ToString();
                            break;
                    }
                }
            }
        }

        private void LoadDataTableToDGV1(DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0) return;
            foreach (DataRow dr in dt.Rows)
            {
                int irowindex = dataGridView3.Rows.Add(dr);
                foreach (DataGridViewColumn col in dataGridView3.Columns)
                {
                    switch (col.Name)
                    {
                        case "col_boxnumberS":
                            dataGridView3.Rows[irowindex].Cells["col_boxnumberS"].Value = dr["boxid"].ToString();
                            break;
                        case "col_palletnumberS":
                            dataGridView3.Rows[irowindex].Cells["col_palletnumberS"].Value = dr["palletid"].ToString();
                            break;
                        case "col_productidS":
                            dataGridView3.Rows[irowindex].Cells["col_productidS"].Value = dr["productid"].ToString();
                            break;
                        case "col_locationS":
                            dataGridView3.Rows[irowindex].Cells["col_locationS"].Value = dr["_location"].ToString();
                            break;
                        case "col_date_inS":
                            dataGridView3.Rows[irowindex].Cells["col_date_inS"].Value = dr["date_in"].ToString();
                            break;
                        case "col_last_operS":
                            dataGridView3.Rows[irowindex].Cells["col_last_operS"].Value = dr["last_oper"].ToString();
                            break;
                        case "col_countS":
                            dataGridView3.Rows[irowindex].Cells["col_countS"].Value = dr["_count"].ToString();
                            break;
                        case "col_commentS":
                            dataGridView3.Rows[irowindex].Cells["col_commentS"].Value = dr["comment"].ToString();
                            break;
                    }
                }
            }
        }

        private void btn_clearsearch_Click(object sender, EventArgs e)
        {
            clearTabpageOut(); 
        }

        private void clearTabpageOut() 
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                    comboBox2.SelectedIndex = -1;
                    searchtextBox.Text = ""; 
                    dataGridView2.Rows.Clear();
                    listBox1.Items.Clear();

                }));
            }
            else
            {
                comboBox2.SelectedIndex = -1;
                searchtextBox.Text = "";
                dataGridView2.Rows.Clear();
                listBox1.Items.Clear();
                searchtextBox.Focus();
            }
        }

        private void clearTabpageOut1()
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                    comboBox4.SelectedIndex = -1;
                    SearchNotOuttextBox.Text = "";
                    dataGridView3.Rows.Clear();
                    listBox1.Items.Clear();

                }));
            }
            else
            {
                comboBox4.SelectedIndex = -1;
                SearchNotOuttextBox.Text = "";
                dataGridView3.Rows.Clear();
                SearchNotOuttextBox.Focus();
            }
        }

        private void clearTabpageOut2()
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                    comboBox1.SelectedIndex = -1;
                    searchinformationtextBox.Text = "";
                    dataGridView1.Rows.Clear();

                }));
            }
            else
            {
                comboBox1.SelectedIndex = -1;
                searchinformationtextBox.Text = "";
                dataGridView1.Rows.Clear();
                searchinformationtextBox.Focus();
            }
        }

        private void clearTabpageOut3()
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                    comboBox5.SelectedIndex = -1;
                    comboBox6.SelectedIndex = -1;
                    ChangeSearchtextBox.Text = "";
                    ChangetextBox.Text = "";
                    dataGridView5.Rows.Clear();

                }));
            }
            else
            {
                comboBox1.SelectedIndex = -1;
                comboBox6.SelectedIndex = -1;
                ChangeSearchtextBox.Text = "";
                ChangetextBox.Text = "";
                dataGridView5.Rows.Clear();
                ChangeSearchtextBox.Focus();
            }
        }

        private void dgv_search_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != 0) return;
            if (e.RowIndex < 0) return; 
            dataGridView2.Rows[e.RowIndex].Cells[0].Value = !(bool)dataGridView2.Rows[e.RowIndex].Cells[0].Value; 
        }

        private void dgv_change_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != 0) return;
            if (e.RowIndex < 0) return;
            dataGridView5.Rows[e.RowIndex].Cells[0].Value = !(bool)dataGridView5.Rows[e.RowIndex].Cells[0].Value;
            foreach (DataGridViewRow dr in dataGridView5.Rows)
            {
                if ((bool)dr.Cells[0].Value == true)
                {
                    string strboxid = dr.Cells["change_col_boxnumber"].Value.ToString();
                    if (!listBox4.Items.Contains(strboxid))
                    {
                        listBox4.Items.Add(strboxid);
                    }
                }
                else if ((bool)dr.Cells[0].Value == false)
                {
                    string strboxid = dr.Cells["change_col_boxnumber"].Value.ToString();
                    if (listBox4.Items.Contains(strboxid))
                    {
                        listBox4.Items.Remove(strboxid);
                    }
                }
            }
        }

        private void btn_add2buffer_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dr in dataGridView2.Rows) 
            {
                if ((bool)dr.Cells[0].Value == true) 
                {
                    string strboxid = dr.Cells["col_boxnumber"].Value.ToString();
                    if (!listBox1.Items.Contains(strboxid)) 
                    {
                        listBox1.Items.Add(strboxid);
                    }
                }
            }
        }

        private void btn_delbufferitem_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex == -1) return;
            var items = listBox1.SelectedItems;
            for (int i = items.Count - 1; i >= 0; i--) 
            {
                listBox1.Items.Remove(items[i]); 
            } 
        }

        private void dgv_search_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex != 0) return;
            if (dataGridView2.Rows.Count > 0) 
            {
                bool bstate = (bool)dataGridView2.Rows[0].Cells["col_checked"].Value; 
                foreach (DataGridViewRow dgvr in dataGridView2.Rows) 
                {
                    dgvr.Cells["col_checked"].Value = !bstate; 
                }
            }
        }

        private void dgv_change_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex != 0) return;
            if (dataGridView5.Rows.Count > 0)
            {
                bool bstate = (bool)dataGridView5.Rows[0].Cells["change_col_checked"].Value;
                foreach (DataGridViewRow dgvr in dataGridView5.Rows)
                {
                    dgvr.Cells["change_col_checked"].Value = !bstate;
                }
            }
            foreach (DataGridViewRow dr in dataGridView5.Rows)
            {
                if ((bool)dr.Cells[0].Value == true)
                {
                    string strboxid = dr.Cells["change_col_boxnumber"].Value.ToString();
                    if (!listBox4.Items.Contains(strboxid))
                    {
                        listBox4.Items.Add(strboxid);
                    }
                }
                else if ((bool)dr.Cells[0].Value == false)
                {
                    string strboxid = dr.Cells["change_col_boxnumber"].Value.ToString();
                    if (listBox4.Items.Contains(strboxid))
                    {
                        listBox4.Items.Remove(strboxid);
                    }
                }
            }
        }

        private void btn_out_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count == 0) return;
            string strboxs = "";
            foreach (object o in listBox1.Items) 
            {
                strboxs += o.ToString() + "\r\n"; 
            }
            string msg = "選取物料箱：\r\n"+ strboxs + "是否確認領出？";
            if (MessageBox.Show(msg, "確認領出", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            if (!_connection.Ping()) _connection.Open(); ;
            if (_connection.Ping()) 
            {
                string strCommand = "";
                foreach (object o in listBox1.Items) 
                {
                    strCommand += string.Format("UPDATE items SET last_oper = '{0}', date_out = '{1}' WHERE boxid = '{2}' and date_out is null;", 
                        usertextBox.Text, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), o.ToString()); 
                }
                MySqlCommand cmd = new MySqlCommand(strCommand, _connection);
                try
                {
                    cmd.ExecuteNonQuery();
                    _clsLog.Info(String.Format("{0}, 領料, {1}", usertextBox.Text, strboxs.Replace("\r\n", "; ")));
                    _thd_outOK = new Thread(BoxOutOK);
                    _thd_outOK.IsBackground = true;
                    _thd_outOK.Start();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                MessageBox.Show("資料庫連線異常，請重新嘗試。", "資料庫異常", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        
        private void BoxOutOK()
        {
            setLabelVisible(label13, true);
            clearTabpageOut();
            Thread.Sleep(3000);
            setLabelVisible(label13, false);
        }

        private void BoxOutNG()
        {
            setLabelVisible(label13, true);
            clearTabpageOut();
            Thread.Sleep(3000);
            setLabelVisible(label13, false);
        }

        public bool IsVaildInput(string word)
        {
            Regex NumandEG = new Regex("[^A-Za-z0-9-_]");
            return !NumandEG.IsMatch(word);
        }

        public bool IsVaildInputEN(string word)
        {
            Regex NumandEG = new Regex("^[BT][0-9]");
            return !NumandEG.IsMatch(word);
        }

        public bool IsVaildInputPEN(string word)
        {
            Regex NumandEG = new Regex("^[P][0-9]");
            return !NumandEG.IsMatch(word);
        }

        public bool IsVaildInputLEN(string word)
        {
            Regex NumandEG = new Regex("^[L]");
            return !NumandEG.IsMatch(word);
        }

        public bool IsVaildInputNo(string word)
        {
            Regex NumandEG = new Regex("^[0-9]*$");
            return !NumandEG.IsMatch(word);
        }

        public string getPalletLocation(string strpalletnumber) 
        {
            if (!_connection.Ping()) _connection.Open();
            if (_connection.Ping()) 
            {
                MySqlCommand cmd = _connection.CreateCommand();
                cmd.CommandText = "select _location from items where palletid = '" + strpalletnumber + "' LIMIT 1";
                MySqlDataReader reader  = cmd.ExecuteReader();

                if (reader.Read())
                {
                    string str_loacaiton = reader.GetString(0);
                    reader.Close();
                    return str_loacaiton;
                }
                else 
                {
                    reader.Close();
                    return ""; 
                }
            }
            return ""; 
        }

        public string getPalletname(string strproductnumber)
        {
            if (!_connection.Ping()) _connection.Open();
            if (_connection.Ping())
            {
                MySqlCommand cmd = _connection.CreateCommand();
                cmd.CommandText = "select palletid from items where productid = '" + strproductnumber + "' LIMIT 1";
                MySqlDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    string strpallet = reader.GetString(0);
                    reader.Close();
                    return strpallet;
                }
                else
                {
                    reader.Close();
                    return "";
                }
            }
            return "";
        }

        public string getcountnumber(string strcountnumber)
        {
            if (!_connection.Ping()) _connection.Open();
            if (_connection.Ping())
            {
                MySqlCommand cmd = _connection.CreateCommand();
                cmd.CommandText = "select _count from items where productid = '" + strcountnumber + "' LIMIT 1";
                MySqlDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    string strcount = reader.GetString(0);
                    reader.Close();
                    return strcount;
                }
                else
                {
                    reader.Close();
                    return "";
                }
            }
            return "";
        }

        public string getPalletProduct(string strpalletnumber)
        {
            if (!_connection.Ping()) _connection.Open();
            if (_connection.Ping())
            {
                MySqlCommand cmd = _connection.CreateCommand();
                cmd.CommandText = "select productid from items where palletid = '" + strpalletnumber + "' LIMIT 1";
                MySqlDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    string strproduct = reader.GetString(0);
                    reader.Close();
                    return strproduct;
                }
                else
                {
                    reader.Close();
                    return "";
                }
            }
            return "";
        }

        private int GetSQLCount(MySqlConnection conn, string strCol, string strVal)
        {
            int ireturn = 0;
            if (!conn.Ping()) conn.Open();
            if (conn.Ping())
            {
                string strcommand = string.Format("select count(*) from items where {0} = '{1}'", strCol, strVal);
                MySqlCommand cmd = new MySqlCommand(strcommand, conn);
                MySqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    for (int i = 0; i < reader.VisibleFieldCount; i++)
                    {
                        if (reader.IsDBNull(i)) continue;
                        switch (reader.GetName(i))
                        {
                            case "count(*)":
                                ireturn = reader.GetInt32(i);
                                break;
                        }
                    }
                }
                reader.Close();
            }
            return ireturn;
        }

        private int GetSQLResultCount(MySqlConnection conn, string strCol, string strVal)
        {
            int ireturn = 0;
            if (!conn.Ping()) conn.Open();
            if (conn.Ping())
            {
                string strcommand = string.Format("select count(*) from items where {0} = '{1}' and date_out is null", strCol, strVal);
                MySqlCommand cmd = new MySqlCommand(strcommand, conn);
                MySqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    for (int i = 0; i < reader.VisibleFieldCount; i++)
                    {
                        if (reader.IsDBNull(i)) continue;
                        switch (reader.GetName(i))
                        {
                            case "count(*)":
                                ireturn = reader.GetInt32(i);
                                break;
                        }
                    }
                }
                reader.Close();
            }
            return ireturn;
        }

        private int GetSQLResultCountAll(MySqlConnection conn)
        {
            int ireturn = 0;
            if (!conn.Ping()) conn.Open();
            if (conn.Ping())
            {
                string strcommand = string.Format("select count(*) from items where date_out is null");
                MySqlCommand cmd = new MySqlCommand(strcommand, conn);
                MySqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    for (int i = 0; i < reader.VisibleFieldCount; i++)
                    {
                        if (reader.IsDBNull(i)) continue;
                        switch (reader.GetName(i))
                        {
                            case "count(*)":
                                ireturn = reader.GetInt32(i);
                                break;
                        }
                    }
                }
                reader.Close();
            }
            return ireturn;
        }

        private int GetSQLResultCountAllTrayBox_in(MySqlConnection conn,string strUser,string strBox)
        {
            int ireturn = 0;
            if (!conn.Ping()) conn.Open();
            if (conn.Ping())
            {
                string strcommand = string.Format("select count(*) from items where last_oper = '{0}' and boxid = '{1}' and  date_out is null",strUser,strBox);
                MySqlCommand cmd = new MySqlCommand(strcommand, conn);
                MySqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    for (int i = 0; i < reader.VisibleFieldCount; i++)
                    {
                        if (reader.IsDBNull(i)) continue;
                        switch (reader.GetName(i))
                        {
                            case "count(*)":
                                ireturn = reader.GetInt32(i);
                                break;
                        }
                    }
                }
                reader.Close();
            }
            return ireturn;
        }

        private int GetSQLResultCountAllTrayBox_out(MySqlConnection conn, string strUser, string strBox)
        {
            int ireturn = 0;
            if (!conn.Ping()) conn.Open();
            if (conn.Ping())
            {
                string strcommand = string.Format("select count(*) from items where last_oper = '{0}' and boxid = '{1}'", strUser, strBox);
                MySqlCommand cmd = new MySqlCommand(strcommand, conn);
                MySqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    for (int i = 0; i < reader.VisibleFieldCount; i++)
                    {
                        if (reader.IsDBNull(i)) continue;
                        switch (reader.GetName(i))
                        {
                            case "count(*)":
                                ireturn = reader.GetInt32(i);
                                break;
                        }
                    }
                }
                reader.Close();
            }
            return ireturn;
        }

        private int GetSQLResultCountReason(MySqlConnection conn, string strKey, string strVal)
        {
            int ireturn = 0;
            if (!conn.Ping()) conn.Open();
            if (conn.Ping())
            {
                string strcommand = string.Format("select count(*) from items where palletid = '{0}' and productid = '{1}' and date_out is null", strKey, strVal);
                MySqlCommand cmd = new MySqlCommand(strcommand, conn);
                MySqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    for (int i = 0; i < reader.VisibleFieldCount; i++)
                    {
                        if (reader.IsDBNull(i)) continue;
                        switch (reader.GetName(i))
                        {
                            case "count(*)":
                                ireturn = reader.GetInt32(i);
                                break;
                        }
                    }
                }
                reader.Close();
            }
            return ireturn;
        }

        private int GetLineSQLCount(MySqlConnection conn, string strCol, string strVal)
        {
            int ireturn = 0;
            if (!conn.Ping()) conn.Open();
            if (conn.Ping())
            {
                string strcommand = string.Format("select count(*) from products where {0} = '{1}'", strCol, strVal);
                MySqlCommand cmd = new MySqlCommand(strcommand, conn);
                MySqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    for (int i = 0; i < reader.VisibleFieldCount; i++)
                    {
                        if (reader.IsDBNull(i)) continue;
                        switch (reader.GetName(i))
                        {
                            case "count(*)":
                                ireturn = reader.GetInt32(i);
                                break;
                        }
                    }
                }
                reader.Close();
            }
            return ireturn;
        }

        private void setLabelVisible(Label clslabel, bool isvisible)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                    clslabel.Visible = isvisible;
                }));
            }
            else
            {
                clslabel.Visible = isvisible;
            }
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            
            DataTable dt_buffer = new DataTable();

            if (comboBox1.SelectedIndex == -1 || searchinformationtextBox.Text == "")
            {
                if (!_connection.Ping()) _connection.Open();
                if (_connection.Ping())
                {
                    string strCommand = "select * from items";
                    MySqlDataAdapter adapter = new MySqlDataAdapter(strCommand, _connection);
                    adapter.Fill(dt_buffer);
                }
            }
            else
            {
                string strkey = "";
                switch (comboBox1.Text.ToUpper())
                {
                    case "PRODUCT ID":
                        strkey = "productid";
                        break;
                    case "BOX NUMBER":
                        strkey = "boxid";
                        break;
                    case "PALLET NUMBER":
                        strkey = "palletid";
                        break;
                }

                string strval = searchinformationtextBox.Text;

                if (!_connection.Ping()) _connection.Open();
                if (_connection.Ping())
                {
                    string strCommand = string.Format("select * from items where {0} = '{1}'", strkey, strval);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(strCommand, _connection);
                    adapter.Fill(dt_buffer);
                }
            }
            LoadDataTableToDGV_search(dt_buffer);
            searchinformationtextBox.Text = "";
        }

        private void LoadDataTableToDGV_search(DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0) return;
            foreach (DataRow dr in dt.Rows)
            {
                int irowindex = dataGridView1.Rows.Add(dr);
                dataGridView1.Rows[irowindex].Cells["Column1"].Value = dr["boxid"].ToString();
                dataGridView1.Rows[irowindex].Cells["Column2"].Value = dr["palletid"].ToString();
                dataGridView1.Rows[irowindex].Cells["Column3"].Value = dr["productid"].ToString();
                dataGridView1.Rows[irowindex].Cells["Column4"].Value = dr["_location"].ToString();
                dataGridView1.Rows[irowindex].Cells["Column5"].Value = dr["date_in"].ToString();
                dataGridView1.Rows[irowindex].Cells["Column6"].Value = dr["date_out"].ToString();
                dataGridView1.Rows[irowindex].Cells["Column7"].Value = dr["last_oper"].ToString();
                //dataGridView1.Rows[irowindex].Cells["Column8"].Value = dr["last_oper_out"].ToString();
                dataGridView1.Rows[irowindex].Cells["Column9"].Value = dr["_count"].ToString();
                dataGridView1.Rows[irowindex].Cells["Column8"].Value = dr["comment"].ToString();
            }
        }

        private void LoadDataTableToDGV_Line(DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0) return;
            foreach (DataRow dr in dt.Rows)
            {
                int irowindex = dataGridView4.Rows.Add(dr);
                dataGridView4.Rows[irowindex].Cells["line"].Value = dr["linename"].ToString();
                dataGridView4.Rows[irowindex].Cells["lineproduct"].Value = dr["lineproductid"].ToString();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1) return; 
            searchinformationtextBox.Focus();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1_Click(null, null);
            }
        }

        private void keyintextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string strboxs = keyintextBox.Text;
                string stroper = usertextBox.Text;
                string strproduct = productIDtextBox.Text;
                string strpallet = palletIDtextBox.Text;
                string strlocation = palletlocationtextBox.Text;
                string strcount = counttextBox.Text;
                string strcomment = commenttextBox.Text;
                listBox1.Items.Add(strboxs);
                if (strboxs == "")
                {
                    MessageBox.Show("資料不完整，請重新輸入。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //_clsLog.Info(usertextBox.Text + ", Output Error Pallet ..., " +  "," + keyintextBox.Text + ", "+  productIDtextBox.Text +", " + palletIDtextBox.Text + "," );
                    _clsLog.Info(String.Format("{0}, Output Error Pallet ... , {1}, {2}, {3}, {4}, {5} ,{6}", stroper, strboxs, strproduct, strpallet, strlocation, strcount, strcomment));
                    return;
                }
                if (!_connection.Ping()) _connection.Open(); ;
                if (_connection.Ping())
                {
                    string strCommand = "";
                    foreach (object o in listBox1.Items)
                    {
                        strCommand += string.Format("UPDATE items SET last_oper = '{0}', date_out = '{1}' WHERE boxid = '{2}' and date_out is null;",
                            usertextBox.Text, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), o.ToString());

                    }
                    MySqlCommand cmd = new MySqlCommand(strCommand, _connection);
                    try
                    {
                        cmd.ExecuteNonQuery();

                        _clsLog.Info(String.Format("{0}, 領料, {1}", usertextBox.Text, strboxs.Replace("\r\n", "; ")));
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("資料庫連線異常，請重新嘗試。", "資料庫異常", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
               
                ShowTrayBoxOlabel.Text = strboxs;
                if (this.listBox2.Items.Count > 0)
                {
                    if (!listBox2.Items.Contains(strboxs))
                    {
                        listBox2.Items.Add(strboxs);
                    }
                }
                else
                {
                    listBox2.Items.Add(strboxs);
                }
                    
                totallabel_out.Text = listBox2.Items.Count.ToString();
                keyintextBox.Text = "";
                keyintextBox.Focus();

                _thd_outOK = new Thread(BoxOutOK);
                _thd_outOK.IsBackground = true;
                _thd_outOK.Start();
            }
        }

        private void interbutton_Click(object sender, EventArgs e)
        {
            string strpallet = palletsettextBox.Text;
            string stroper = usertextBox.Text;

            if (radioButton1.Checked == true)
                palletlabel01.Text = palletsettextBox.Text;
            else if(radioButton2.Checked == true)
                palletlabel02.Text = palletsettextBox.Text;
            else if (radioButton3.Checked == true)
                palletlabel03.Text = palletsettextBox.Text;
            else if (radioButton4.Checked == true)
                palletlabel04.Text = palletsettextBox.Text;
            else if (radioButton5.Checked == true)
                palletlabel05.Text = palletsettextBox.Text;
            else if (radioButton6.Checked == true)
                palletlabel06.Text = palletsettextBox.Text;
            else if (radioButton7.Checked == true)
                palletlabel07.Text = palletsettextBox.Text;
            else if (radioButton8.Checked == true)
                palletlabel08.Text = palletsettextBox.Text;
            else if (radioButton9.Checked == true)
                palletlabel09.Text = palletsettextBox.Text;
            else if (radioButton10.Checked == true)
                palletlabel10.Text = palletsettextBox.Text;
            else if (radioButton11.Checked == true)
                palletlabel11.Text = palletsettextBox.Text;
            else if (radioButton12.Checked == true)
                palletlabel12.Text = palletsettextBox.Text;
            else if (radioButton13.Checked == true)
                palletlabel13.Text = palletsettextBox.Text;
            else if (radioButton14.Checked == true)
                palletlabel14.Text = palletsettextBox.Text;
            else if (radioButton15.Checked == true)
                palletlabel15.Text = palletsettextBox.Text;
            else if (radioButton16.Checked == true)
                palletlabel16.Text = palletsettextBox.Text;

            palletsettextBox.Text = "";
        }

        private void SetInOK()
        {
            setLabelVisible(label35, true);
            clearTabpageSet();
            Thread.Sleep(3000);
            setLabelVisible(label35, false);
        }

        private void StartOK()
        {
            setLabelVisible(label17, true);
            clearTabpageSet();
            Thread.Sleep(3000);
            setLabelVisible(label17, false);
        }

        private void clearTabpageSet()
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                    palletsettextBox.Text = "";
                    //palletsettextBox.Text = "";
                    palletsettextBox.Focus();

                }));
            }
            else
            {
                palletsettextBox.Text = "";
                palletsettextBox.Focus();
            }
        }

        private void tabcontrol_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabcontrol.SelectedIndex == 0)
            {
                productIDtextBox.Focus();

            }
            else if (tabcontrol.SelectedIndex == 1)
            {
                keyintextBox.Focus();
            }
            else if(tabcontrol.SelectedIndex == 2)
            {
                searchinformationtextBox.Focus();
            }
            else if (tabcontrol.SelectedIndex == 3)
            {
                palletsettextBox.Focus();
            }
            else if (tabcontrol.SelectedIndex == 4)
            {
                palletsettextBox.Focus();
            }
            else if (tabcontrol.SelectedIndex == 5)
            {
                CountResulelabel.Text = "";
                setsearchbutton_Click(null, null);
                palletsettextBox.Focus();
            }
        }

        private void setsearchbutton_Click(object sender, EventArgs e)
        {

        }

        private void productsettextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                interbutton_Click(null, null);
            }
        }

        private void palletsettextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                interbutton_Click(null, null);
            }
        }

        public static bool dataGridViewToCSV(DataGridView dataGridView)
        {
            if (dataGridView.Rows.Count == 0)
            {
                MessageBox.Show("沒有數據可導出!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "CSV檔(*.csv)|*.csv | 文字檔(*.txt)| *.txt | All files (*.*)|*.*";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.CreatePrompt = true;
            saveFileDialog.FileName = null;
            saveFileDialog.Title = "保存";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                Stream stream = saveFileDialog.OpenFile();
                StreamWriter sw = new StreamWriter(stream, System.Text.Encoding.GetEncoding(-0));
                string strLine = "";
                try
                {
                    //表頭
                    for (int i = 0; i < dataGridView.ColumnCount; i++)
                    {
                        if (i > 0)
                            strLine += ",";
                        strLine += dataGridView.Columns[i].HeaderText;
                    }
                    strLine.Remove(strLine.Length - 1);
                    sw.WriteLine(strLine);
                    strLine = "";
                    //表的內容
                    for (int j = 0; j < dataGridView.Rows.Count; j++)
                    {
                        strLine = "";
                        int colCount = dataGridView.Columns.Count;
                        for (int k = 0; k < colCount; k++)
                        {
                            if (k > 0 && k < colCount)
                                strLine += ",";
                            if (dataGridView.Rows[j].Cells[k].Value == null)
                                strLine += "";
                            else
                            {
                                string cell = dataGridView.Rows[j].Cells[k].Value.ToString().Trim();
                                //防止裡面含有特殊符號
                                cell = cell.Replace("\"", "\"\"");
                                cell = "\"" + cell + "\"";
                                strLine += cell;
                            }
                        }
                        sw.WriteLine(strLine);
                    }
                    sw.Close();
                    stream.Close();
                    MessageBox.Show("數據被導出到：" + saveFileDialog.FileName.ToString(), "導出完畢", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "導出錯誤", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }
            }
            return true;
        }

        private void Savefilebutton_Click(object sender, EventArgs e)
        {
            dataGridViewToCSV(dataGridView1);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            dataGridViewToCSV(dataGridView2);
        }

        private void counttextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                trayboxIDtextBox.Focus();
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        private void listBox15_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            MessageBox.Show("hello world");
        }

        private void lineproductbutton_Click(object sender, EventArgs e)
        {
            string strlineproduct = producttextBox.Text;
            if (comboBox3.SelectedIndex == -1 || strlineproduct == "" )
            {
                MessageBox.Show("資料不完整，請重新輸入。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                string strlinekey = "";
                switch (comboBox3.Text.ToUpper())
                {
                    case "401":
                        strlinekey = "01";
                        break;
                    case "402":
                        strlinekey = "02";
                        break;
                    case "403":
                        strlinekey = "03";
                        break;
                    case "404":
                        strlinekey = "04";
                        break;
                    case "405":
                        strlinekey = "05";
                        break;
                    case "406":
                        strlinekey = "06";
                        break;
                    case "407":
                        strlinekey = "07";
                        break;
                    case "408":
                        strlinekey = "08";
                        break;
                    case "409":
                        strlinekey = "09";
                        break;
                    case "410":
                        strlinekey = "10";
                        break;
                    case "411":
                        strlinekey = "11";
                        break;
                    case "412":
                        strlinekey = "12";
                        break;
                    case "413":
                        strlinekey = "13";
                        break;
                    case "414":
                        strlinekey = "14";
                        break;
                    case "415":
                        strlinekey = "15";
                        break;
                }

                if (_connection.Ping())
                {
                    MySqlCommand cmd = _connection.CreateCommand();

                    int icount = GetLineSQLCount(_connection, "linename", strlinekey);

                    if (icount > 0)
                    {
                        cmd.CommandText = string.Format("UPDATE products SET lineproductid = '{0}' WHERE linename = {1};", strlineproduct, strlinekey);
                    }
                    else
                    {
                        cmd.CommandText = string.Format("INSERT INTO products ( linename, lineproductid) VALUES('{0}','{1}');", strlinekey, strlineproduct);
                    }

                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("資料庫連線異常，請重新嘗試。", "資料庫異常", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            button10_Click(null, null);
        }

        private void palletIDtextBox_TextChanged(object sender, EventArgs e)
        {
            if (IsVaildInput(palletIDtextBox.Text))
            {
                string strproductnumber = productIDtextBox.Text;
                palletIDtextBox.Text = getPalletname(strproductnumber);
                string strpalletnumber = palletIDtextBox.Text;
                palletlocationtextBox.Text = getPalletLocation(strpalletnumber);
                string strcountnumber = productIDtextBox.Text;
                counttextBox.Text = getcountnumber(strcountnumber);
            }
            else return;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void startbutton_Click(object sender, EventArgs e)
        {
            string[] strpalletnumber = new string[16];
            strpalletnumber[0] = palletlabel01.Text;
            strpalletnumber[1] = palletlabel02.Text;
            strpalletnumber[2] = palletlabel03.Text;
            strpalletnumber[3] = palletlabel04.Text;
            strpalletnumber[4] = palletlabel05.Text;
            strpalletnumber[5] = palletlabel06.Text;
            strpalletnumber[6] = palletlabel07.Text;
            strpalletnumber[7] = palletlabel08.Text;
            strpalletnumber[8] = palletlabel09.Text;
            strpalletnumber[9] = palletlabel10.Text;
            strpalletnumber[10] = palletlabel11.Text;
            strpalletnumber[11] = palletlabel12.Text;
            strpalletnumber[12] = palletlabel13.Text;
            strpalletnumber[13] = palletlabel14.Text;
            strpalletnumber[14] = palletlabel15.Text;
            strpalletnumber[15] = palletlabel16.Text;
            PalletdataGridView1.Rows.Clear();
            DataTable dt_buffer = new DataTable();
            DataTable dt_buffer1 = new DataTable();
            DataTable dt_buffer2 = new DataTable();
            DataTable dt_buffer3 = new DataTable();
            DataTable dt_buffer4 = new DataTable();
            DataTable dt_buffer5 = new DataTable();
            DataTable dt_buffer6 = new DataTable();
            DataTable dt_buffer7 = new DataTable();
            DataTable dt_buffer8 = new DataTable();
            DataTable dt_buffer9 = new DataTable();
            DataTable dt_buffer10 = new DataTable();
            DataTable dt_buffer11 = new DataTable();
            DataTable dt_buffer12 = new DataTable();
            DataTable dt_buffer13 = new DataTable();
            DataTable dt_buffer14 = new DataTable();
            DataTable dt_buffer15 = new DataTable();
            DataTable dt_buffer16 = new DataTable();
            string[,] strvalue = new string[16,4];
            int[,] strcount = new int[16, 4];
            for (int j = 0; j < 16; j++)
            {
                if (!_connection.Ping()) _connection.Open();
                if (_connection.Ping())
                {
                    string strCommand = string.Format("select distinct productid from items where palletid ='{0}' and date_out is null", strpalletnumber[j]);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(strCommand, _connection);
                    adapter.Fill(dt_buffer);
                }

                if (dt_buffer.Rows.Count == 0)
                    return;
                else if (dt_buffer.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_buffer.Rows.Count; i++)
                    {
                        strvalue[j, i] = dt_buffer.Rows[i][0].ToString(); 
                    }
                }
            }
            
            LoadDataTableToProduct1(dt_buffer);
            LoadDataTableToProductA(dt_buffer, PalletdataGridView1);
            LoadDataTableToProductA(dt_buffer, PalletdataGridView2);
            for (int i = 0;i<16;i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    strcount[i,j] = GetSQLResultCountReason(_connection, strpalletnumber[i], strvalue[i, j]);
                }
            }

            Plabel11.Text = strcount[0,0].ToString();
            Plabel12.Text = strcount[0,1].ToString();
            Plabel13.Text = strcount[0,2].ToString();
            Plabel14.Text = strcount[0,3].ToString();
            Plabel21.Text = strcount[1,0].ToString();
            Plabel22.Text = strcount[1,1].ToString();
            Plabel23.Text = strcount[1,2].ToString();
            Plabel24.Text = strcount[1,3].ToString();
            Plabel31.Text = strcount[2, 0].ToString();
            Plabel32.Text = strcount[2, 1].ToString();
            Plabel33.Text = strcount[2, 2].ToString();
            Plabel34.Text = strcount[2, 3].ToString();
            Plabel41.Text = strcount[3, 0].ToString();
            Plabel42.Text = strcount[3, 1].ToString();
            Plabel43.Text = strcount[3, 2].ToString();
            Plabel44.Text = strcount[3, 3].ToString();
            Plabel51.Text = strcount[4, 0].ToString();
            Plabel52.Text = strcount[4, 1].ToString();
            Plabel53.Text = strcount[4, 2].ToString();
            Plabel54.Text = strcount[4, 3].ToString();
            Plabel61.Text = strcount[5, 0].ToString();
            Plabel62.Text = strcount[5, 1].ToString();
            Plabel63.Text = strcount[5, 2].ToString();
            Plabel64.Text = strcount[5, 3].ToString();
            Plabel71.Text = strcount[6, 0].ToString();
            Plabel72.Text = strcount[6, 1].ToString();
            Plabel73.Text = strcount[6, 2].ToString();
            Plabel74.Text = strcount[6, 3].ToString();
            Plabel81.Text = strcount[7, 0].ToString();
            Plabel82.Text = strcount[7, 1].ToString();
            Plabel83.Text = strcount[7, 2].ToString();
            Plabel84.Text = strcount[7, 3].ToString();
            Plabel91.Text = strcount[8, 0].ToString();
            Plabel92.Text = strcount[8, 1].ToString();
            Plabel93.Text = strcount[8, 2].ToString();
            Plabel94.Text = strcount[8, 3].ToString();
            Plabel101.Text = strcount[9, 0].ToString();
            Plabel102.Text = strcount[9, 1].ToString();
            Plabel103.Text = strcount[9, 2].ToString();
            Plabel104.Text = strcount[9, 3].ToString();
            Plabel111.Text = strcount[10, 0].ToString();
            Plabel112.Text = strcount[10, 1].ToString();
            Plabel113.Text = strcount[10, 2].ToString();
            Plabel114.Text = strcount[10, 3].ToString();
            Plabel121.Text = strcount[11, 0].ToString();
            Plabel122.Text = strcount[11, 1].ToString();
            Plabel123.Text = strcount[11, 2].ToString();
            Plabel124.Text = strcount[11, 3].ToString();
            Plabel131.Text = strcount[12, 0].ToString();
            Plabel132.Text = strcount[12, 1].ToString();
            Plabel133.Text = strcount[12, 2].ToString();
            Plabel134.Text = strcount[12, 3].ToString();
            Plabel141.Text = strcount[13, 0].ToString();
            Plabel142.Text = strcount[13, 1].ToString();
            Plabel143.Text = strcount[13, 2].ToString();
            Plabel144.Text = strcount[13, 3].ToString();
            Plabel151.Text = strcount[14, 0].ToString();
            Plabel152.Text = strcount[14, 1].ToString();
            Plabel153.Text = strcount[14, 2].ToString();
            Plabel154.Text = strcount[14, 3].ToString();
            Plabel161.Text = strcount[15, 0].ToString();
            Plabel162.Text = strcount[15, 1].ToString();
            Plabel163.Text = strcount[15, 2].ToString();
            Plabel164.Text = strcount[15, 3].ToString();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void SearchNotOutbutton_Click(object sender, EventArgs e)
        {
            int strcount;
            dataGridView3.Rows.Clear();
            DataTable dt_buffer = new DataTable();

            if (comboBox4.SelectedIndex == -1 || SearchNotOuttextBox.Text == "")
            {
                if (!_connection.Ping()) _connection.Open();
                if (_connection.Ping())
                {
                    string strCommand = "select * from items where date_out is null";
                    MySqlDataAdapter adapter = new MySqlDataAdapter(strCommand, _connection);
                    adapter.Fill(dt_buffer);
                    strcount = GetSQLResultCountAll(_connection);
                    CountResulelabel.Text = strcount.ToString();
                }
            }
            else
            {
                string strkey = "";
                switch (comboBox4.Text.ToUpper())
                {
                    case "PRODUCT ID":
                        strkey = "productid";
                        break;
                    case "BOX NUMBER":
                        strkey = "boxid";
                        break;
                    case "PALLET NUMBER":
                        strkey = "palletid";
                        break;
                }

                string strval = SearchNotOuttextBox.Text;

                if (!_connection.Ping()) _connection.Open();
                if (_connection.Ping())
                {
                    string strCommand = string.Format("select * from items where {0} = '{1}' and date_out is null", strkey, strval);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(strCommand, _connection);
                    adapter.Fill(dt_buffer);
                    strcount = GetSQLResultCount(_connection, strkey, strval);
                    CountResulelabel.Text = strcount.ToString();
                }
            }
            LoadDataTableToDGV1(dt_buffer);
            SearchNotOuttextBox.Text = "";
        }

        private void SearchInClearbutton_Click(object sender, EventArgs e)
        {
            clearTabpageOut1();
        }

        private void SearchInListbutton_Click(object sender, EventArgs e)
        {
            dataGridViewToCSV(dataGridView3);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            dataGridView4.Rows.Clear();

            DataTable dt_buffer = new DataTable();

            if (!_connection.Ping()) _connection.Open();
            if (_connection.Ping())
            {
                string strCommand = "select * from products";
                MySqlDataAdapter adapter = new MySqlDataAdapter(strCommand, _connection);
                adapter.Fill(dt_buffer);
            }
            LoadDataTableToDGV_Line(dt_buffer);
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex == -1) return;
            searchinformationtextBox.Focus();
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.SelectedIndex == -1) return;
            searchinformationtextBox.Focus();
        }


        private void LoadDataTableToProduct1(DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0) return;
            foreach (DataRow dr in dt.Rows)
            {
                int irowindex = PalletdataGridView1.Rows.Add(dr);
                foreach (DataGridViewColumn col in PalletdataGridView1.Columns)
                {
                    switch (col.Name)
                    {
                        case "product1":
                            PalletdataGridView1.Rows[irowindex].Cells["product1"].Value = dr["productid"].ToString();
                            break;
                    }
                }
            }
        }

        private void LoadDataTableToProductA(DataTable dt , DataGridView dataGridViewA)
        {
            dataGridViewA.Rows.Clear();
            if (dt == null || dt.Rows.Count == 0) return;
            foreach (DataRow dr in dt.Rows)
            {
                int irowindex = dataGridViewA.Rows.Add(dr);
                foreach (DataGridViewColumn col in dataGridViewA.Columns)
                {
                    switch (col.Name)
                    {
                        case "product1":
                            dataGridViewA.Rows[irowindex].Cells["product1"].Value = dr["productid"].ToString();
                            break;
                        case "product2":
                            dataGridViewA.Rows[irowindex].Cells["product2"].Value = dr["productid"].ToString();
                            break;
                        case "product3":
                            dataGridViewA.Rows[irowindex].Cells["product3"].Value = dr["productid"].ToString();
                            break;
                        case "product4":
                            dataGridViewA.Rows[irowindex].Cells["product4"].Value = dr["productid"].ToString();
                            break;
                        case "product5":
                            dataGridViewA.Rows[irowindex].Cells["product5"].Value = dr["productid"].ToString();
                            break;
                        case "product6":
                            dataGridViewA.Rows[irowindex].Cells["product6"].Value = dr["productid"].ToString();
                            break;
                        case "product7":
                            dataGridViewA.Rows[irowindex].Cells["product7"].Value = dr["productid"].ToString();
                            break;
                        case "product8":
                            dataGridViewA.Rows[irowindex].Cells["product8"].Value = dr["productid"].ToString();
                            break;
                        case "product9":
                            dataGridViewA.Rows[irowindex].Cells["product9"].Value = dr["productid"].ToString();
                            break;
                        case "product10":
                            dataGridViewA.Rows[irowindex].Cells["product10"].Value = dr["productid"].ToString();
                            break;
                        case "product11":
                            dataGridViewA.Rows[irowindex].Cells["product11"].Value = dr["productid"].ToString();
                            break;
                        case "product12":
                            dataGridViewA.Rows[irowindex].Cells["product12"].Value = dr["productid"].ToString();
                            break;
                        case "product13":
                            dataGridViewA.Rows[irowindex].Cells["product13"].Value = dr["productid"].ToString();
                            break;
                        case "product14":
                            dataGridViewA.Rows[irowindex].Cells["product14"].Value = dr["productid"].ToString();
                            break;
                        case "product15":
                            dataGridViewA.Rows[irowindex].Cells["product15"].Value = dr["productid"].ToString();
                            break;
                        case "product16":
                            dataGridViewA.Rows[irowindex].Cells["product16"].Value = dr["productid"].ToString();
                            break;
                    }
                }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            string[] strpalletnumber = new string[16];
            strpalletnumber[0] = palletlabel01.Text;
            strpalletnumber[1] = palletlabel02.Text;
            strpalletnumber[2] = palletlabel03.Text;
            strpalletnumber[3] = palletlabel04.Text;
            strpalletnumber[4] = palletlabel05.Text;
            strpalletnumber[5] = palletlabel06.Text;
            strpalletnumber[6] = palletlabel07.Text;
            strpalletnumber[7] = palletlabel08.Text;
            strpalletnumber[8] = palletlabel09.Text;
            strpalletnumber[9] = palletlabel10.Text;
            strpalletnumber[10] = palletlabel11.Text;
            strpalletnumber[11] = palletlabel12.Text;
            strpalletnumber[12] = palletlabel13.Text;
            strpalletnumber[13] = palletlabel14.Text;
            strpalletnumber[14] = palletlabel15.Text;
            strpalletnumber[15] = palletlabel16.Text;
           
            DataTable[] dt_bufferA = new DataTable[16];
            DataTable dt_buffer = new DataTable();
            DataTable dt_buffer1 = new DataTable();
            DataTable dt_buffer2 = new DataTable();
            DataTable dt_buffer3 = new DataTable();
            DataTable dt_buffer4 = new DataTable();
            DataTable dt_buffer5 = new DataTable();
            DataTable dt_buffer6 = new DataTable();
            DataTable dt_buffer7 = new DataTable();
            DataTable dt_buffer8 = new DataTable();
            DataTable dt_buffer9 = new DataTable();
            DataTable dt_buffer10 = new DataTable();
            DataTable dt_buffer11 = new DataTable();
            DataTable dt_buffer12 = new DataTable();
            DataTable dt_buffer13 = new DataTable();
            DataTable dt_buffer14 = new DataTable();
            DataTable dt_buffer15 = new DataTable();
            DataTable dt_buffer16 = new DataTable();

            DataGridView[] PalletdataGrid = new DataGridView[16];
            PalletdataGrid[0] = PalletdataGridView1;
            PalletdataGrid[1] = PalletdataGridView2;
            PalletdataGrid[2] = PalletdataGridView3;
            PalletdataGrid[3] = PalletdataGridView4;
            PalletdataGrid[4] = PalletdataGridView5;
            PalletdataGrid[5] = PalletdataGridView6;
            PalletdataGrid[6] = PalletdataGridView7;
            PalletdataGrid[7] = PalletdataGridView8;
            PalletdataGrid[8] = PalletdataGridView9;
            PalletdataGrid[9] = PalletdataGridView10;
            PalletdataGrid[10] = PalletdataGridView11;
            PalletdataGrid[11] = PalletdataGridView12;
            PalletdataGrid[12] = PalletdataGridView13;
            PalletdataGrid[13] = PalletdataGridView14;
            PalletdataGrid[14] = PalletdataGridView15;
            PalletdataGrid[15] = PalletdataGridView16;
            for (int j = 0; j < 16; j++)
            {
                PalletdataGrid[j].Rows.Clear();
            }
            string[,] strvalue = new string[16, 30];
            int[,] strcount = new int[16, 30];
            for (int j = 0; j < 16; j++)
            {
                dt_buffer.Clear();
                if (!_connection.Ping()) _connection.Open();
                if (_connection.Ping())
                {
                    string strCommand = string.Format("select distinct productid from items where palletid ='{0}' and date_out is null", strpalletnumber[j]);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(strCommand, _connection);
                    adapter.Fill(dt_buffer);
                }

                if (dt_buffer.Rows.Count == 0)
                    continue;
                else if (dt_buffer.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_buffer.Rows.Count; i++)
                    {
                        strvalue[j, i] = dt_buffer.Rows[i][0].ToString();
                    }
                }

                for (int i = 0; i < dt_buffer.Rows.Count; i++)
                {
                    strcount[j, i] = GetSQLResultCountReason(_connection, strpalletnumber[j], strvalue[j, i]);
                }
                LoadDataTableToProductA(dt_buffer, PalletdataGrid[j]);
                Plabel11.Text = strcount[0, 0].ToString();
                Plabel12.Text = strcount[0, 1].ToString();
                Plabel13.Text = strcount[0, 2].ToString();
                Plabel14.Text = strcount[0, 3].ToString();
                Plabel21.Text = strcount[1, 0].ToString();
                Plabel22.Text = strcount[1, 1].ToString();
                Plabel23.Text = strcount[1, 2].ToString();
                Plabel24.Text = strcount[1, 3].ToString();
                Plabel31.Text = strcount[2, 0].ToString();
                Plabel32.Text = strcount[2, 1].ToString();
                Plabel33.Text = strcount[2, 2].ToString();
                Plabel34.Text = strcount[2, 3].ToString();
                Plabel41.Text = strcount[3, 0].ToString();
                Plabel42.Text = strcount[3, 1].ToString();
                Plabel43.Text = strcount[3, 2].ToString();
                Plabel44.Text = strcount[3, 3].ToString();
                Plabel51.Text = strcount[4, 0].ToString();
                Plabel52.Text = strcount[4, 1].ToString();
                Plabel53.Text = strcount[4, 2].ToString();
                Plabel54.Text = strcount[4, 3].ToString();
                Plabel61.Text = strcount[5, 0].ToString();
                Plabel62.Text = strcount[5, 1].ToString();
                Plabel63.Text = strcount[5, 2].ToString();
                Plabel64.Text = strcount[5, 3].ToString();
                Plabel71.Text = strcount[6, 0].ToString();
                Plabel72.Text = strcount[6, 1].ToString();
                Plabel73.Text = strcount[6, 2].ToString();
                Plabel74.Text = strcount[6, 3].ToString();
                Plabel81.Text = strcount[7, 0].ToString();
                Plabel82.Text = strcount[7, 1].ToString();
                Plabel83.Text = strcount[7, 2].ToString();
                Plabel84.Text = strcount[7, 3].ToString();
                Plabel91.Text = strcount[8, 0].ToString();
                Plabel92.Text = strcount[8, 1].ToString();
                Plabel93.Text = strcount[8, 2].ToString();
                Plabel94.Text = strcount[8, 3].ToString();
                Plabel101.Text = strcount[9, 0].ToString();
                Plabel102.Text = strcount[9, 1].ToString();
                Plabel103.Text = strcount[9, 2].ToString();
                Plabel104.Text = strcount[9, 3].ToString();
                Plabel111.Text = strcount[10, 0].ToString();
                Plabel112.Text = strcount[10, 1].ToString();
                Plabel113.Text = strcount[10, 2].ToString();
                Plabel114.Text = strcount[10, 3].ToString();
                Plabel121.Text = strcount[11, 0].ToString();
                Plabel122.Text = strcount[11, 1].ToString();
                Plabel123.Text = strcount[11, 2].ToString();
                Plabel124.Text = strcount[11, 3].ToString();
                Plabel131.Text = strcount[12, 0].ToString();
                Plabel132.Text = strcount[12, 1].ToString();
                Plabel133.Text = strcount[12, 2].ToString();
                Plabel134.Text = strcount[12, 3].ToString();
                Plabel141.Text = strcount[13, 0].ToString();
                Plabel142.Text = strcount[13, 1].ToString();
                Plabel143.Text = strcount[13, 2].ToString();
                Plabel144.Text = strcount[13, 3].ToString();
                Plabel151.Text = strcount[14, 0].ToString();
                Plabel152.Text = strcount[14, 1].ToString();
                Plabel153.Text = strcount[14, 2].ToString();
                Plabel154.Text = strcount[14, 3].ToString();
                Plabel161.Text = strcount[15, 0].ToString();
                Plabel162.Text = strcount[15, 1].ToString();
                Plabel163.Text = strcount[15, 2].ToString();
                Plabel164.Text = strcount[15, 3].ToString();
            }
            
        }

        private void SearchNotOuttextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                SearchNotOutbutton_Click(null, null);
        }

        private void PalletdataGridView_DoubleClick(object sender, EventArgs e)
        {
            radioButton1.Checked = true;
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            dataGridView5.Rows.Clear();
            DataTable dt_buffer = new DataTable();

            if (comboBox5.SelectedIndex == -1 || ChangeSearchtextBox.Text == "")
            {
                if (!_connection.Ping()) _connection.Open();
                if (_connection.Ping())
                {
                    string strCommand = "select * from items where date_out is null";
                    MySqlDataAdapter adapter = new MySqlDataAdapter(strCommand, _connection);
                    adapter.Fill(dt_buffer);
                }
            }
            else
            {
                string strkey = "";
                switch (comboBox5.Text.ToUpper())
                {
                    case "PRODUCT ID":
                        strkey = "productid";
                        break;
                    case "BOX NUMBER":
                        strkey = "boxid";
                        break;
                    case "PALLET NUMBER":
                        strkey = "palletid";
                        break;
                }

                string strval = ChangeSearchtextBox.Text;

                if (!_connection.Ping()) _connection.Open();
                if (_connection.Ping())
                {
                    string strCommand = string.Format("select * from items where {0} = '{1}' and date_out is null", strkey, strval);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(strCommand, _connection);
                    adapter.Fill(dt_buffer);
                }
            }
            LoadDataTableToDGVChange(dt_buffer);
        }

        private void ChangeSearchtextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                button14_Click(null, null);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            string strvaluekey = "";
            switch (comboBox5.Text.ToUpper())
            {
                case "PRODUCT ID":
                    strvaluekey = "productid";
                    break;
                case "BOX NUMBER":
                    strvaluekey = "boxid";
                    break;
                case "PALLET NUMBER":
                    strvaluekey = "palletid";
                    break;
            }
            string strboxs = "";
            foreach (object o in listBox4.Items)
            {
                strboxs += o.ToString() + "\r\n";
            }
            string strchangekey = "";
            string strchangekeypallet = ChangeSearchtextBox.Text;
            switch (comboBox6.Text.ToUpper())
            {
                case "PRODUCT ID":
                    strchangekey = "productid";
                    break;
                case "PALLET NUMBER":
                    strchangekey = "palletid";
                    break;
            }
            string strchangeval = ChangetextBox.Text;

            if (!_connection.Ping()) _connection.Open();
            if (_connection.Ping())
            {
                string strCommand = "";
                foreach (object o in listBox4.Items)
                {
                    strCommand += string.Format("UPDATE items SET {0} = '{1}' WHERE boxid = '{2}' and {3} = '{4}' and date_out is null;",
                        strchangekey, strchangeval, o.ToString(),strvaluekey, strchangekeypallet);
                }
                MySqlCommand cmd = new MySqlCommand(strCommand, _connection);
                try
                {
                    cmd.ExecuteNonQuery();
                    _clsLog.Info(String.Format("{0}, 變更, {1}, {2}, {3}", usertextBox.Text, strboxs.Replace("\r\n", "; "), ChangeSearchtextBox.Text, ChangetextBox.Text));
                    _thd_outOK = new Thread(BoxOutOK);
                    _thd_outOK.IsBackground = true;
                    _thd_outOK.Start();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            listBox4.Items.Clear();
            ChangetextBox.Text = "";
            button14_Click(null, null);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            listBox4.Items.Clear();
            clearTabpageOut3();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            totallabel_in.Text = "";
        }

        private void button17_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            totallabel_out.Text = "";
        }

        private void ChangetextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                button16_Click(null, null);
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
            clearTabpageOut2();
        }
    }
}
