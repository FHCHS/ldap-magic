using System;
using System.Collections;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Web;
using System.Windows.Forms;
using Google.GData.Apps;
using Google.GData.Client;
using WindowsApplication1.utils;


// CONSIDERATIONS
// This program is designed around SQL server 2000 and Active Directory 2003, SQLserver 2005 uses different methods for accessing databases table information and column information this will need to upgraded when the move is made
// This program depends on the existence of the columns picked in the save data, if columns names change the mappings will have to be remapped
// This program depends on the use of # temporary tables and therefore needs read & write capabilities to the selected database
// This program attempts to check for empty lists and handle them robustly however there is a high chance if problem will occurs it is from a mishandeld empty list passed between functions
// This program uses paging on expected large results from AD if too large of a pull (greater than 1499 records) comes from a  AD query it will break a non-paged function resulting in failed transactions
// The default SQL connection is created so timeouts will match the default selected (not sure but Im gussing it is ~30s) indefinitely held connections could be a problem while the program remains open
// There is a large set of created objects and memory leaks may abound. Most objects are closed however some rely on the garbage collector
// Many functions in tools are overloaded and when updating a tool be sure to be careful to make sure each overloaded method will match the new standards
// This program has a logfile which will tell which function failed when an exception is thrown and will attempt to add useful data ( timestamp, funciton failed, passed variables)
// This program requires the SQL server to have access to AD this ADSI stored procedure links the two
// This program must be run with a user with sufficient rights to read and write to AD users and groups
//
// CLASSES
// Each of the classes below
// hold the data that represents the fields the corresponding tab
// has a to dictionary method for making a text file savable set of strings
// has a load funciton which takes a dicitonary
// groupSynch
// userSynch
// userStateChange
// executionOrder
//
// toolset
// contains many functions which could be pulled out of the main code some could be added from the UI buttons
//
// objectADSqlsyncGroup
// is an excuse not to put everything inside of an on_click event for a button
// 
// UI CODE
// UI events HAVE CODE IN THEM WHICH IS NOT ABSTRACTED INTO FUNCTIONS AND WILL NEED INDIVIDUAL ATTNETION WHEN UPDATING
// it is arranged by tab and then grouped by the groupbox in which it appears on the form
// UI events set the values in the datastructure like classes
// UI events call execution methods



// have any exceptions write to a log file

namespace WindowsApplication1
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }
        // create objects to hold save data
        GroupSynch groupconfig = new GroupSynch();
        UserSynch userconfig = new UserSynch();
		GmailUsers guserconfig = new GmailUsers();
        executionOrder execution = new executionOrder();
        UserStateChange usermapping = new UserStateChange();                          
        ToolSet tools = new ToolSet();
        LogFile log = new LogFile();
        ObjectADSqlsyncGroup groupSyncr = new ObjectADSqlsyncGroup();
		ObjectADGoogleSync gmailSyncr = new ObjectADGoogleSync();


        // UI DIALOG  DATA ENTRY EVENTS FOR USER MAPPING TAB
        private void usersMap_mapping_description_TextChanged(object sender, EventArgs e)
        {
            userconfig.Notes = users_mapping_description.Text.ToString();
        }
        private void usersMap_user_Table_View_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (userconfig.DBCatalog != "" && userconfig.DataServer != "")
            {
                //populates table dialog with tables or views depending on the results of a query
                ArrayList tableList = new ArrayList();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + userconfig.DataServer.ToString() + ";Initial Catalog=" + userconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");
                SqlCommand sqlComm;
                sqlConn.Open();
                // create the command object
                if (users_user_Table_View.Text.ToLower() == "table")
                {
                    sqlComm = new SqlCommand("SELECT name FROM sysobjects where TYPE = 'U' order by NAME", sqlConn);
                }
                else
                {
                    sqlComm = new SqlCommand("select name from SYSOBJECTS where TYPE = 'V' order by NAME", sqlConn);
                }
                SqlDataReader r = sqlComm.ExecuteReader();
                while (r.Read())
                {
                    tableList.Add((string)r[0].ToString().Trim());
                }
                r.Close();
                sqlConn.Close();

                userconfig.User_table_view = users_user_Table_View.Text.ToString();
                users_user_source.DataSource = tableList;
            }
            else
            {
                MessageBox.Show("Please set the dataserver and catalog");
            }
        }
        private void usersMap_user_source_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (userconfig.DBCatalog != "" && userconfig.DataServer != "")
            {
                //populates columns dialog with columns depending on the results of a query
                ArrayList columnList = new ArrayList();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + userconfig.DataServer.ToString() + ";Initial Catalog=" + userconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");

                sqlConn.Open();
                // create the command object
                SqlCommand sqlComm = new SqlCommand("SELECT column_name FROM information_schema.columns WHERE table_name = '" + users_user_source.Text.ToString() + "'", sqlConn);
                SqlDataReader r = sqlComm.ExecuteReader();
                while (r.Read())
                {
                    columnList.Add((string)r[0].ToString().Trim());
                }
                r.Close();
                sqlConn.Close();

                userconfig.User_dbTable = users_user_source.Text.ToString();
                users_user_Lname.DataSource = columnList;
                users_user_sAMAccountName.DataSource = columnList.Clone();
            }
            else
            {
                MessageBox.Show("Please select table or view");
            }
        }
        private void usersMap_user_CN_SelectedIndexChanged(object sender, EventArgs e)
        {
            userconfig.User_Lname = users_user_Lname.Text.ToString();
        }
        private void usersMap_user_sAMAccountName_SelectedIndexChanged(object sender, EventArgs e)
        {
            userconfig.User_sAMAccount = users_user_sAMAccountName.Text.ToString();
        }
        private void usersMap_user_where_TextChanged(object sender, EventArgs e)
        {
            userconfig.User_where = users_user_where.Text.ToString();
        }
        private void usersMap_baseUserOU_TextChanged(object sender, EventArgs e)
        {
            userconfig.BaseUserOU = users_baseUserOU.Text.ToString();
        }



        // UI DIALOG  DATA ENTRY EVENTS FOR USER SYNCH TAB

        private void users_user_Table_View_SelectedIndexChanged(object sender, EventArgs e)
        {
            // only valid for SQL server 2000
            if (userconfig.DBCatalog != "" && userconfig.DataServer != "")
            {
                //populates table dialog with tables or views depending on the results of a query
                ArrayList tableList = new ArrayList();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + userconfig.DataServer.ToString() + ";Initial Catalog=" + userconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");
                SqlCommand sqlComm;
                sqlConn.Open();
                // create the command object
                if (users_user_Table_View.Text.ToLower() == "table")
                {
                    sqlComm = new SqlCommand("SELECT name FROM sysobjects where TYPE = 'U' order by NAME", sqlConn);
                }
                else
                {
                    sqlComm = new SqlCommand("select name from SYSOBJECTS where TYPE = 'V' order by NAME", sqlConn);
                }
                SqlDataReader r = sqlComm.ExecuteReader();
                while (r.Read())
                {
                    tableList.Add((string)r[0].ToString().Trim());
                }
                r.Close();
                sqlConn.Close();

                userconfig.User_table_view = users_user_Table_View.Text.ToString();
                users_user_source.DataSource = tableList;
            }
            else
            {
                MessageBox.Show("Please set the dataserver and catalog");
            }
        }
        private void users_user_source_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (userconfig.DBCatalog != "" && userconfig.DataServer != "")
            {
                //populates columns dialog with columns depending on the results of a query
                ArrayList columnList = new ArrayList();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + userconfig.DataServer.ToString() + ";Initial Catalog=" + userconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");

                sqlConn.Open();
                // create the command object
                SqlCommand sqlComm = new SqlCommand("SELECT column_name FROM information_schema.columns WHERE table_name = '" + users_user_source.Text.ToString() + "'", sqlConn);
                SqlDataReader r = sqlComm.ExecuteReader();
                while (r.Read())
                {
                    columnList.Add((string)r[0].ToString().Trim());
                }
                r.Close();
                sqlConn.Close();

                userconfig.User_dbTable = users_user_source.Text.ToString();
                users_user_Lname.DataSource = columnList;
                users_user_Fname.DataSource = columnList.Clone();
                users_user_city.DataSource = columnList.Clone();
                users_user_State.DataSource = columnList.Clone();
                users_user_PostalCode.DataSource = columnList.Clone();
                users_user_Address.DataSource = columnList.Clone();
                users_user_Mobile.DataSource = columnList.Clone();
                users_user_sAMAccountName.DataSource = columnList.Clone();
                users_user_password.DataSource = columnList.Clone();
//                users_user_email.DataSource = columnList.Clone();
                ADColumn.DataSource = tools.ADobjectAttribute();
                columnList.Add("Static Value");
                SQLColumn.DataSource = columnList.Clone();

            }
            else
            {
                MessageBox.Show("Please select table or view");
            }
        }

        private void users_user_Lname_SelectedIndexChanged(object sender, EventArgs e)
        {
            userconfig.User_Lname = users_user_Lname.Text.ToString();
        }
        private void users_user_Fname_SelectedIndexChanged(object sender, EventArgs e)
        {
            userconfig.User_Fname = users_user_Fname.Text.ToString();
        }
        private void users_user_city_SelectedIndexChanged(object sender, EventArgs e)
        {
            userconfig.User_city = users_user_city.Text.ToString();
        }
        private void users_user_State_SelectedIndexChanged(object sender, EventArgs e)
        {
            userconfig.User_State = users_user_State.Text.ToString();
        }
        private void users_user_PostalCode_SelectedIndexChanged(object sender, EventArgs e)
        {
            userconfig.User_Zip = users_user_PostalCode.Text.ToString();
        }
        private void users_user_Address_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            userconfig.User_Address = users_user_Address.Text.ToString();
        }
        private void users_user_Mobile_SelectedIndexChanged(object sender, EventArgs e)
        {
            userconfig.User_Mobile = users_user_Mobile.Text.ToString();
        }
        private void users_user_sAMAccountName_SelectedIndexChanged(object sender, EventArgs e)
        {
            userconfig.User_sAMAccount = users_user_sAMAccountName.Text.ToString();
        }
        private void users_user_password_SelectedIndexChanged(object sender, EventArgs e)
        {
            userconfig.User_password = users_user_password.Text.ToString();
        }
        //private void users_user_email_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    userconfig.User_mail = users_user_email.Text.ToString();
        //}

        private void users_user_where_TextChanged(object sender, EventArgs e)
        {
            userconfig.User_where = users_user_where.Text.ToString();
        }
        private void users_baseUserOU_Leave(object sender, EventArgs e)
        {
            if (users_baseUserOU.Text.ToString() != "")
            {
                if (tools.Exists(users_baseUserOU.Text.ToString()))
                {
                    userconfig.BaseUserOU = users_baseUserOU.Text.ToString();
                }
                else
                {

                    DialogResult button = MessageBox.Show("OU LDAP://" + users_baseUserOU.Text.ToString() + " does Not exist shall I create it", "Nonexistent OU", MessageBoxButtons.YesNo);

                    if (button == DialogResult.Yes)
                    {
                        tools.CreateOURecursive(users_baseUserOU.Text.ToString(), log);
                        userconfig.BaseUserOU = users_baseUserOU.Text.ToString();
                    }

                    if (button == DialogResult.No)
                    {
                        users_baseUserOU.Focus();
                    }
                }
            }
        }
        private void users_holdingTank_Leave(object sender, EventArgs e)
        {
            if (users_holdingTank.Text.ToString() != "")
            {
                if (tools.Exists(users_holdingTank.Text.ToString()))
                {
                    userconfig.UserHoldingTank = users_holdingTank.Text.ToString();
                }
                else
                {

                    DialogResult button = MessageBox.Show("OU LDAP://" + users_holdingTank.Text.ToString() + " does Not exist shall I create it", "Nonexistent OU", MessageBoxButtons.YesNo);

                    if (button == DialogResult.Yes)
                    {
                        tools.CreateOURecursive(users_holdingTank.Text.ToString(), log);
                        userconfig.UserHoldingTank = users_holdingTank.Text.ToString();
                    }

                    if (button == DialogResult.No)
                    {
                        users_holdingTank.Focus();
                    }
                }
            }
        }
        private void users_group_TextChanged(object sender, EventArgs e)
        {
            userconfig.UniversalGroup = users_group.Text.ToString();
        }   
        private void users_mapping_description_TextChanged(object sender, EventArgs e)
        {
            userconfig.Notes = users_mapping_description.Text.ToString();
        }
        // BUTTONS FOR THE TAB 
        private void users_Save_button(object sender, EventArgs e)
        {
            // MessageBox.Show(mike.Columns[0] + " is the name of a column");

            StringBuilder CustomsString = new StringBuilder();
            Dictionary<string, string> properties = new Dictionary<string, string>();

            int i = 0;
            int j = 0;

            for (i = 0; i < mappinggrid.RowCount - 1; i++)
            {
                for (j = 0; j < mappinggrid.ColumnCount; j++)
                {
                    if (mappinggrid.Rows[i].Cells[j].Value == null)
                    {
                        CustomsString.Append("^");
                    }
                    else
                    {
                        CustomsString.Append(mappinggrid.Rows[i].Cells[j].Value.ToString() + "^");
                    }
                }
                CustomsString.Remove(CustomsString.Length - 1, 1);
                CustomsString.Append("&");
            }
            if (CustomsString.Length != 0)
            {
                CustomsString.Remove(CustomsString.Length - 1, 1);
            }
            userconfig.CustomsString = CustomsString.ToString();
            

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // create a file stream, where "c:\\testing.txt" is the file path
                System.IO.FileStream fs = new System.IO.FileStream(saveFileDialog1.FileName, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write, System.IO.FileShare.ReadWrite);

                // create a stream writer
                System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.ASCII);

                // write to file (buffer), where textbox1 is your text box
                properties = userconfig.ToDictionary();
                ICollection<string> c = properties.Keys;

                foreach (string str in c)
                {
                    sw.WriteLine("{0} | {1:C}", str, properties[str]);
                }

                // flush buffer (so the text really goes into the file)
                sw.Flush();

                // close stream writer and file
                sw.Close();
                fs.Close();
            }
        }        
        private void users_Execute_Click(object sender, EventArgs e)
        {
			int i;
			StopWatch timer = new StopWatch();
			timer.Start();
			groupSyncr.ExecuteUserSync(userconfig, tools, log);
			timer.Stop();
			MessageBox.Show("bulk " + timer.GetElapsedTimeSecs().ToString());
			StringBuilder result = new StringBuilder();
			StringBuilder result2 = new StringBuilder();
			result.Append("***************************\n*                         *\n*        Transactions     *\n*                         *\n***************************");
			for (i = 0; i < log.transactions.Count; i++)
			{
				result.Append(log.transactions[i].ToString() + "\n");
			}

			result.Append("***************************\n*                         *\n*        Warnings         *\n*                         *\n***************************");

			for (i = 0; i < log.warnings.Count; i++)
			{
				result.Append(log.warnings[i].ToString() + "\n");
			}

			result.Append("***************************\n*                         *\n*        Errors           *\n*                         *\n***************************");
			for (i = 0; i < log.errors.Count; i++)
			{
				result2.Append(log.errors[i].ToString() + "\n");
			}

			// save log to disk
			SaveFileDialog saveFileDialog1 = new SaveFileDialog();
			saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
			saveFileDialog1.FilterIndex = 2;
			saveFileDialog1.RestoreDirectory = true;
			if (saveFileDialog1.ShowDialog() == DialogResult.OK)
			{
				// create a file stream, where "c:\\testing.txt" is the file path
				System.IO.FileStream fs = new System.IO.FileStream(saveFileDialog1.FileName, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write, System.IO.FileShare.ReadWrite);

				// create a stream writer
				System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.ASCII);

				// write to file (buffer), where textbox1 is your text box
				sw.WriteLine("{0}", result2);
				sw.WriteLine("{0}", result);


				// flush buffer (so the text really goes into the file)
				sw.Flush();

				// close stream writer and file
				sw.Close();
				fs.Close();
			}
			//            group_result1.AppendText(result.ToString());
			//            group_result2.AppendText(result2.ToString());
        }
        private void users_open_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> properties = new Dictionary<string, string>();
            DataTable customs = new DataTable();
            BindingSource bs = new BindingSource();

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                StreamReader re = File.OpenText(openFileDialog1.FileName);

                string input = null;
                while ((input = re.ReadLine()) != null)
                {
                    string[] parts = input.Split('|');
                    properties.Add(parts[0].Trim(), parts[1].Trim());
                }
                re.Close();
            }

            // Load values into text boxes
            // reload properties each time as they are overwritten with the combo object trigger events
            userconfig.load(properties);
            DBserver.Text = userconfig.DataServer;
            userconfig.load(properties);
            Catalog.Text = userconfig.DBCatalog;
            userconfig.load(properties);

            users_user_Table_View.Text = userconfig.User_table_view;
            userconfig.load(properties);
            users_user_source.Text = userconfig.User_dbTable;
            userconfig.load(properties);
            users_user_sAMAccountName.Text = userconfig.User_sAMAccount;
            userconfig.load(properties);
            users_user_Lname.Text = userconfig.User_Lname;
            userconfig.load(properties);
            users_user_Fname.Text = userconfig.User_Fname;
            userconfig.load(properties);
            users_user_city.Text = userconfig.User_city;
            userconfig.load(properties);
            users_user_Mobile.Text = userconfig.User_Mobile;
            userconfig.load(properties);
            users_user_State.Text = userconfig.User_State;
            userconfig.load(properties);
            users_user_Address.Text = userconfig.User_Address;
            userconfig.load(properties);
            users_user_password.Text = userconfig.User_password;
            userconfig.load(properties);
            users_user_PostalCode.Text = userconfig.User_Zip;
            userconfig.load(properties);
            users_user_where.Text = userconfig.User_where;
            userconfig.load(properties);
            users_holdingTank.Text = userconfig.UserHoldingTank;
            userconfig.load(properties);
            users_mapping_description.Text = userconfig.Notes;
            users_baseUserOU.Text = userconfig.BaseUserOU;
            users_group.Text = userconfig.UniversalGroup;
            userconfig.load(properties);
            //users_user_emailDomain.Text = userconfig.UserEmailDomain;
            //userconfig.load(properties);

            // clear the grid in case the previous open user file had values in it
            mappinggrid.Rows.Clear();
            DataGridViewRow row = new DataGridViewRow();
            DataGridViewCell cell0;
            DataGridViewCell cell1;
            DataGridViewCell cell2;
            ArrayList ad = tools.ADobjectAttribute();
            ArrayList sql = tools.SqlColumns(userconfig);

            // fill the grid with values
            for (int i = 0; i < userconfig.UserCustoms.Rows.Count; i++)
            {
                cell0 = new DataGridViewComboBoxCell();
                cell1 = new DataGridViewComboBoxCell();
                cell2 = new DataGridViewTextBoxCell();
                cell0.Value = userconfig.UserCustoms.Rows[i][0].ToString();
                cell1.Value = userconfig.UserCustoms.Rows[i][1].ToString();
                cell2.Value = userconfig.UserCustoms.Rows[i][2].ToString();
                row.Cells.Add(cell0);
                row.Cells.Add(cell1);
                row.Cells.Add(cell2);
                mappinggrid.Rows.Add(row);

                row = new DataGridViewRow();
            }

            // create and set the comboboxes
            sql.Add("Static Value");  
            SQLColumn.DataSource = sql;
            ADColumn.DataSource = ad;
            // auto resize must be off to keep data errors from happening
            SQLColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            ADColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
               
        }



        /* UI DIALOG  DATA ENTRY EVENTS FOR GROUP SYNCH TAB */
        private void group_group_Table_View_SelectedIndexChanged(object sender, EventArgs e)
        {
            // only valid for SQL server 2000
            if (groupconfig.DBCatalog != "" && groupconfig.DataServer != "")
            {
                //populates table dialog with tables or views depending on the results of a query
                ArrayList tableList = new ArrayList();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + groupconfig.DataServer.ToString() + ";Initial Catalog=" + groupconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");
                SqlCommand sqlComm;
                sqlConn.Open();
                // create the command object
                if (group_group_Table_View.Text.ToLower() == "table")
                {
                    sqlComm = new SqlCommand("SELECT name FROM sysobjects where TYPE = 'U' order by NAME", sqlConn);
                }
                else
                {
                    sqlComm = new SqlCommand("SELECT name from SYSOBJECTS where TYPE = 'V' order by NAME", sqlConn);
                }
                SqlDataReader r = sqlComm.ExecuteReader();
                while (r.Read())
                {
                    tableList.Add((string)r[0].ToString().Trim());
                }
                r.Close();
                sqlConn.Close();

                groupconfig.Group_table_view = group_group_Table_View.Text.ToString();
                group_group_source.DataSource = tableList;
            }
            else
            {
                MessageBox.Show("Please set the dataserver and catalog");
            }
        }
        private void group_group_source_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (groupconfig.DBCatalog != "" && groupconfig.DataServer != "")
            {
                //populates columns dialog with columns depending on the results of a query
                ArrayList columnList = new ArrayList();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + groupconfig.DataServer.ToString() + ";Initial Catalog=" + groupconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");

                sqlConn.Open();
                // create the command object
                SqlCommand sqlComm = new SqlCommand("SELECT column_name FROM information_schema.columns WHERE table_name = '" + group_group_source.Text.ToString() + "'", sqlConn);
                SqlDataReader r = sqlComm.ExecuteReader();
                while (r.Read())
                {
                    columnList.Add((string)r[0].ToString().Trim());
                }
                r.Close();
                sqlConn.Close();

                groupconfig.Group_dbTable = group_group_source.Text.ToString();
                group_group_CN.DataSource = columnList;
                group_group_sAMAccountName.DataSource = columnList.Clone();
            }
            else
            {
                MessageBox.Show("Please select table or view");
            }

        }
        private void group_group_CN_SelectedIndexChanged(object sender, EventArgs e)
        {
            groupconfig.Group_CN = group_group_CN.Text.ToString();
        }
        private void group_group_sAMAccount_SelectedIndexChanged(object sender, EventArgs e)
        {
            groupconfig.Group_sAMAccount = group_group_sAMAccountName.Text.ToString();
        }
        private void group_group_where_TextChanged(object sender, EventArgs e)
        {
            groupconfig.Group_where = group_group_where.Text.ToString();
        }
        private void group_group_prepend_TextChanged(object sender, EventArgs e)
        {
            groupconfig.Group_Append = group_group_prepend.Text.ToString();
        }

        private void group_user_Table_View_SelectedIndexChanged(object sender, EventArgs e)
        {
            // only valid for SQL server 2000
            if (groupconfig.DBCatalog != "" && groupconfig.DataServer != "")
            {
                //populates table dialog with tables or views depending on the results of a query
                ArrayList tableList = new ArrayList();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + groupconfig.DataServer.ToString() + ";Initial Catalog=" + groupconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");
                SqlCommand sqlComm;
                sqlConn.Open();
                // create the command object
                if (group_user_Table_View.Text.ToLower() == "table")
                {
                    sqlComm = new SqlCommand("SELECT name FROM sysobjects where TYPE = 'U' order by NAME", sqlConn);
                }
                else
                {
                    sqlComm = new SqlCommand("select name from SYSOBJECTS where TYPE = 'V' order by NAME", sqlConn);
                }
                SqlDataReader r = sqlComm.ExecuteReader();
                while (r.Read())
                {
                    tableList.Add((string)r[0].ToString().Trim());
                }
                r.Close();
                sqlConn.Close();

                groupconfig.User_table_view = group_user_Table_View.Text.ToString();
                group_user_source.DataSource = tableList;
            }
            else
            {
                MessageBox.Show("Please set the dataserver and catalog");
            }
        }
        private void group_user_source_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (groupconfig.DBCatalog != "" && groupconfig.DataServer != "")
            {
                //populates columns dialog with columns depending on the results of a query
                ArrayList columnList = new ArrayList();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + groupconfig.DataServer.ToString() + ";Initial Catalog=" + groupconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");

                sqlConn.Open();
                // create the command object
                SqlCommand sqlComm = new SqlCommand("SELECT column_name FROM information_schema.columns WHERE table_name = '" + group_user_source.Text.ToString() + "'", sqlConn);
                SqlDataReader r = sqlComm.ExecuteReader();
                while (r.Read())
                {
                    columnList.Add((string)r[0].ToString().Trim());
                }
                r.Close();
                sqlConn.Close();

                groupconfig.User_dbTable = group_user_source.Text.ToString();
                group_user_Group_reference.DataSource = columnList;
                group_user_sAMAccountName.DataSource = columnList.Clone();
            }
            else
            {
                MessageBox.Show("Please select table or view");
            }
        }

        private void group_user_sAMAccountName_SelectedIndexChanged(object sender, EventArgs e)
        {
            groupconfig.User_sAMAccount = group_user_sAMAccountName.Text.ToString();
        }
        private void group_user_where_TextChanged(object sender, EventArgs e)
        {
            groupconfig.User_where = group_user_where.Text.ToString();
        }
        private void group_user_Group_reference_SelectedIndexChanged(object sender, EventArgs e)
        {
            groupconfig.User_Group_Reference = group_user_Group_reference.Text.ToString();
        }
        private void group_baseGroupOU_TextChanged(object sender, EventArgs e)
        {
            groupconfig.BaseGroupOU = group_baseGroupOU.Text.ToString();
        }
        private void group_baseUserOU_TextChanged(object sender, EventArgs e)
        {
            groupconfig.BaseUserOU = group_baseUserOU.Text.ToString();
        }
        private void group_mapping_description_TextChanged(object sender, EventArgs e)
        {
            groupconfig.Notes = group_mapping_description.Text.ToString();
        }
        // BUTTONS FOR THE TAB
        private void group_Save_button(object sender, EventArgs e)
        {
            Dictionary<string, string> properties = new Dictionary<string, string>();
            int i = 0;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // create a file stream, where "c:\\testing.txt" is the file path
                System.IO.FileStream fs = new System.IO.FileStream(saveFileDialog1.FileName, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write, System.IO.FileShare.ReadWrite);

                // create a stream writer
                System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.ASCII);

                // write to file (buffer), where textbox1 is your text box
                properties = groupconfig.ToDictionary();
                ICollection<string> c = properties.Keys;
                i = c.Count;
                foreach (string str in c)
                {
                    sw.WriteLine("{0} | {1:C}", str, properties[str]);
                }

                // flush buffer (so the text really goes into the file)
                sw.Flush();

                // close stream writer and file
                sw.Close();
                fs.Close();
            }
        }
        private void group_open_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> properties = new Dictionary<string, string>();

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                StreamReader re = File.OpenText(openFileDialog1.FileName);

                string input = null;
                while ((input = re.ReadLine()) != null)
                {
                    string[] parts = input.Split('|');
                    properties.Add(parts[0].Trim(), parts[1].Trim());
                }
                re.Close();
            }
            // Load values into text boxes
            // reload properties each time as they are overwritten with the combo object trigger events
            groupconfig.Load(properties);
            DBserver.Text = groupconfig.DataServer;
            Catalog.Text = groupconfig.DBCatalog;
            group_group_Table_View.Text = groupconfig.Group_table_view;
            groupconfig.Load(properties);
            group_group_source.Text = groupconfig.Group_dbTable;
            groupconfig.Load(properties);
            group_group_CN.Text = groupconfig.Group_CN;
            groupconfig.Load(properties);
            group_group_sAMAccountName.Text = groupconfig.Group_sAMAccount;
            groupconfig.Load(properties);
            group_group_where.Text = groupconfig.Group_where;
            groupconfig.Load(properties);
            group_group_prepend.Text = groupconfig.Group_Append;

            group_user_Table_View.Text = groupconfig.User_table_view;
            groupconfig.Load(properties);
            group_user_source.Text = groupconfig.User_dbTable;
            groupconfig.Load(properties);
            group_user_Group_reference.Text = groupconfig.User_Group_Reference;
            groupconfig.Load(properties);
            group_user_sAMAccountName.Text = groupconfig.User_sAMAccount;
            groupconfig.Load(properties);
            group_user_where.Text = groupconfig.User_where;

            group_mapping_description.Text = groupconfig.Notes;
            group_baseUserOU.Text = groupconfig.BaseUserOU;
            group_baseGroupOU.Text = groupconfig.BaseGroupOU;

        }
        private void group_execute_now_Click(object sender, EventArgs e)
		{
			int i;
			StopWatch timer = new StopWatch();
			timer.Start();
			groupSyncr.ExecuteGroupSync(groupconfig, tools, log);
			timer.Stop();
			MessageBox.Show("bulk " + timer.GetElapsedTimeSecs().ToString());

			//string sqlgroupsTable = "#sqltableADTransfertesttoy";
			//SqlConnection sqlConn = new SqlConnection("Data Source=" + groupconfig.DataServer + ";Initial Catalog=" + groupconfig.DBCatalog + ";Integrated Security=SSPI;");
			//SqlCommand sqlComm;
			//sqlConn.Open();
			//sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + groupconfig.Group_sAMAccount + ") AS " + groupconfig.Group_sAMAccount + ", RTRIM(" + groupconfig.Group_CN + ") + '" + groupconfig.Group_Append + "' AS " + groupconfig.Group_CN + " INTO " + sqlgroupsTable + " FROM " + groupconfig.Group_dbTable + " ORDER BY " + groupconfig.Group_CN, sqlConn);
			//sqlComm.ExecuteNonQuery();
			//SqlDataReader debugreader;
			//StringBuilder debug = new StringBuilder();
			//int debugfieldcount;
			//sqlComm = new SqlCommand("select * FROM " + sqlgroupsTable, sqlConn);
			//debugreader = sqlComm.ExecuteReader();
			//debugfieldcount = debugreader.FieldCount;
			//while (debugreader.Read())
			//{
			//    for (i = 0; i < debugfieldcount; i++)
			//    {
			//        debug.Append((string)debugreader[i].ToString() + ",");
			//    }
			//    debug.AppendLine();
			//}
			//debugreader.Close();
			//group_result1.AppendText(debug.ToString());

			StringBuilder result = new StringBuilder();
			StringBuilder result2 = new StringBuilder();
			result.Append("***************************\n*                         *\n*        Transactions     *\n*                         *\n***************************");
			for (i = 0; i < log.transactions.Count; i++)
			{
				result.Append(log.transactions[i].ToString() + "\n");
			}

			result.Append("***************************\n*                         *\n*        Warnings         *\n*                         *\n***************************");

			for (i = 0; i < log.warnings.Count; i++)
			{
				result.Append(log.warnings[i].ToString() + "\n");
			}

			result.Append("***************************\n*                         *\n*        Errors           *\n*                         *\n***************************");
			for (i = 0; i < log.errors.Count; i++)
			{
				result2.Append(log.errors[i].ToString() + "\n");
			}
			group_result1.AppendText(result.ToString());
			group_result2.AppendText(result2.ToString());
			// groupSyncr.execute(groupconfig, tools, log);
			// users_result1.Text log.transactions.ToString();
			// users_result2.Text = log.errors.ToString();
			// MessageBox.Show("compelete");
		}

		//deprecated
		private void group_push_for_virus_Click(object sender, EventArgs e)
		{
			MessageBox.Show(System.Web.HttpUtility.UrlEncode("` ` , [ ] hi / hi  \\ + * ? | = ; :", System.Text.Encoding.GetEncoding("utf-8")));
			//MessageBox.Show(System.Web.HttpUtility.UrlDecode("free%2A+test+%2F%5C%5B%5D%27%3F%3A%3B%7C%3D%2C%2B%3E%3C%22+bob").Replace("+", " "));

			//StopWatch timer = new StopWatch();
			//timer.Start();
			//groupSyncr.execute(groupconfig, tools, log, this);
			//timer.Stop();
			//MessageBox.Show("non bulk " + timer.GetElapsedTimeSecs().ToString());
			//int i;
			//for (i = 0; i < log.transactions.Count; i++)
			//{
			//    group_result1.AppendText(log.transactions[i].ToString() + "\n");
			//}
			//for (i = 0; i < log.errors.Count; i++)
			//{
			//    group_result2.AppendText(log.errors[i].ToString() + "\n");
			//}

			//if (tools.Authenticate("mne4d7", "blah", "OU=Atis,OU=FHCHS,DC=FHCHS,DC=EDU") == true)
			//    group_result1.AppendText("found you");
			//else
			//    group_result1.AppendText("failure you");
			//SqlConnection sqlConn = new SqlConnection("Data Source=fhcsvdb;Initial Catalog=soniswebdatabase;Integrated Security=SSPI;");

			//sqlConn.Open();
			//// create the command object
			//SqlCommand sqlComm = new SqlCommand("SELECT soc_sec, first_name, ssn FROM name WHERE first_name like 'xa'", sqlConn);
			//SqlDataReader r = sqlComm.ExecuteReader();
			//while (r.Read())
			//{
			//    string username = (string)r["first_name"];
			//    string userID = (string)r["soc_sec"];
			//    group_result1.AppendText(username);
			//    group_result1.AppendText(userID);
			//}
			//r.Close();
			//sqlConn.Close();
			//foreach (string abc in tools.EnumerateOU("OU=Atis,OU=FHCHS,DC=FHCHS,DC=EDU"))
			//{
			//    group_result1.AppendText(abc);
			//}
		}
		private void group_see_query_Click(object sender, EventArgs e)
		{
			group_result1.Clear();
			group_result1.AppendText("This is your group query \n");
			group_result1.AppendText("Select ");
			group_result1.AppendText(groupconfig.Group_CN);
			group_result1.AppendText(", ");
			group_result1.AppendText(groupconfig.Group_sAMAccount);
			group_result1.AppendText(" From ");
			group_result1.AppendText(groupconfig.Group_dbTable);
			if (groupconfig.Group_where != string.Empty)
			{
				group_result1.AppendText(" Where ");
				group_result1.AppendText(groupconfig.Group_where);
			}
			group_result1.AppendText("\n");

			group_result2.Clear();
			group_result2.AppendText("This is your user query \n");
			group_result2.AppendText("Select ");
			group_result2.AppendText(groupconfig.User_Group_Reference);
			group_result2.AppendText(", ");
			group_result2.AppendText(groupconfig.User_sAMAccount);
			group_result2.AppendText(" From ");
			group_result2.AppendText(groupconfig.User_dbTable);
			if (groupconfig.User_where != string.Empty)
			{
				group_result2.AppendText(" Where ");
				group_result2.AppendText(groupconfig.User_where);
			}
			group_result2.AppendText("\n");
		}


        // UI DIALOG  DATA ENTRY EVENTS FOR CONFIGURATION TAB
        private void test_data_source_Click(object sender, EventArgs e)
        {
            SqlConnection sqlConn = new SqlConnection("Data Source=" + DBserver.Text.ToString() + ";Initial Catalog=" + Catalog.Text.ToString() + ";Integrated Security=SSPI;");
            try
            {
                sqlConn.Open();
                groupconfig.DBCatalog = Catalog.Text.ToString();
                userconfig.DBCatalog = Catalog.Text.ToString();
                guserconfig.DBCatalog = Catalog.Text.ToString();
                groupconfig.DataServer = DBserver.Text.ToString();
                userconfig.DataServer = DBserver.Text.ToString();
                guserconfig.DataServer = DBserver.Text.ToString();
                sqlConn.Close();
            }
            catch
            {
                MessageBox.Show("Cannot Locate Database");
                Catalog.Text = "";
                DBserver.Text = "";
            }
        }
        // BUTTONS FOR THE TAB
        private void DBserver_TextChanged(object sender, EventArgs e)
        {
            ArrayList tableList = new ArrayList();
            System.Data.SqlClient.SqlConnection SqlCon = new System.Data.SqlClient.SqlConnection("server=" + DBserver.Text.ToString() + ";Integrated Security=SSPI;");
            SqlCon.Open();

            System.Data.SqlClient.SqlCommand SqlCom = new System.Data.SqlClient.SqlCommand();
            SqlCom.Connection = SqlCon;
            SqlCom.CommandType = CommandType.StoredProcedure;
            SqlCom.CommandText = "sp_databases";

            System.Data.SqlClient.SqlDataReader r;
            r = SqlCom.ExecuteReader();

            while (r.Read())
            {
                tableList.Add((string)r[0].ToString().Trim());
            }
            r.Close();
            SqlCon.Close();

            Catalog.DataSource = tableList;

        }


        //UI DIALOG  DATA ENTRY EVENTS FOR USERMAP TAB
        private void usermap_user_source_SelectedIndexChanged(object sender, EventArgs e)
        {
            usermapping.User_dbTable = usermapping_user_source.Text;

            if (userMapping_DBServerName.Text == "")
                userMapping_DBServerName.Focus();
            else if (userMapping_DatabaseName.Text == "")
                userMapping_DatabaseName.Focus();
            else if (usermapping_user_source.Text == "")
                usermapping_user_source.Focus();
            else
            {
                //populates columns dialog with columns depending on the results of a query
                ArrayList columnList = new ArrayList();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + userMapping_DBServerName.Text + ";Initial Catalog=" + userMapping_DatabaseName.Text + ";Integrated Security=SSPI;");

                sqlConn.Open();
                // create the command object
                SqlCommand sqlComm = new SqlCommand("SELECT column_name FROM information_schema.columns WHERE table_name = '" + usermapping_user_source.Text.ToString() + "'", sqlConn);
                SqlDataReader r = sqlComm.ExecuteReader();
                while (r.Read())
                {
                    columnList.Add((string)r[0].ToString().Trim());
                }
                r.Close();
                sqlConn.Close();

                usermapping_user_sAMAccountName.DataSource = columnList;
            }

        }
        private void usermap_user_sAMAccountName_SelectedIndexChanged(object sender, EventArgs e)
        {
            usermapping.User_sAMAccount = usermapping_user_sAMAccountName.Text;
        }
        private void usermap_user_table_view_SelectedIndexChanged(object sender, EventArgs e)
        {
            // only valid for SQL server 2000
            if (userMapping_DBServerName.Text == "")
                userMapping_DBServerName.Focus();
            else if (userMapping_DatabaseName.Text == "")
                userMapping_DatabaseName.Focus();
            else
            {
                //populates table dialog with tables or views depending on the results of a query
                ArrayList tableList = new ArrayList();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + userMapping_DBServerName.Text + ";Initial Catalog=" + userMapping_DatabaseName.Text + ";Integrated Security=SSPI;");
                SqlCommand sqlComm;
                sqlConn.Open();
                // create the command object
                if (usermapping_user_table_view.Text.ToLower() == "table")
                {
                    sqlComm = new SqlCommand("SELECT name FROM sysobjects where TYPE = 'U' order by NAME", sqlConn);
                }
                else
                {
                    sqlComm = new SqlCommand("SELECT name from SYSOBJECTS where TYPE = 'V' order by NAME", sqlConn);
                }
                SqlDataReader r = sqlComm.ExecuteReader();
                while (r.Read())
                {
                    tableList.Add((string)r[0].ToString().Trim());
                }
                r.Close();
                sqlConn.Close();

                usermapping.User_table_view = usermapping_user_table_view.Text;
                usermapping_user_source.DataSource = tableList;
            }
        }
        private void usermap_user_where_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {
            usermapping.User_where = usermapping_user_where.Text;
        }
        private void usermap_OU_OU_TextChanged(object sender, EventArgs e)
        {

        }
        private void usermap_OU_list_groups_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }
        private void usermap_baseOU_TextChanged(object sender, EventArgs e)
        {

        }
        private void usermap_disableDays_TextChanged(object sender, EventArgs e)
        {

        }
        private void usermap_fromOU_TextChanged(object sender, EventArgs e)
        {

        }
        private void usermap_toOU_TextChanged(object sender, EventArgs e)
        {

        }

        //UI FOR LIST SELECTOR
        private void execution_add_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                execution.execution_order.Add(openFileDialog1.FileName.ToString().Clone());
            }
            execution_order_list.DataSource = execution.execution_order;
        }
        private void execution_remove_Click(object sender, EventArgs e)
        {
            if (execution_order_list.SelectedIndex > 0)
            {
                execution.execution_order.RemoveAt(execution.execution_order.IndexOf(execution_order_list.SelectedIndex.ToString()));
            }
        }
        private void userMap_factAdd_Click(object sender, EventArgs e)
        {
            //Note this saves SelectedIndex.  This means that any DB level column reording destroys these mappings.
            if (userMapping_factList.SelectedItems.Count == 1)
            {
                //Edit
                ListViewItem bob = userMapping_factList.SelectedItems[0];

//                if (usermapping_user_table_view.
                bob.Text = userMapping_factName.Text;
                bob.Name = userMapping_factName.Text;
                bob.SubItems["serverName"].Text = userMapping_DBServerName.Text;
                bob.SubItems["serverName"].Tag = userMapping_DBServerName.Text;
                bob.SubItems["databaseIndex"].Text = userMapping_DatabaseName.SelectedItem.ToString();
                bob.SubItems["databaseIndex"].Tag = userMapping_DatabaseName.SelectedIndex;
                bob.SubItems["typeIndex"].Text = usermapping_user_table_view.SelectedItem.ToString();
                bob.SubItems["typeIndex"].Tag = usermapping_user_table_view.SelectedIndex;
                bob.SubItems["tableIndex"].Text = usermapping_user_source.SelectedItem.ToString();
                bob.SubItems["tableIndex"].Tag = usermapping_user_source.SelectedIndex;
                bob.SubItems["columnIndex"].Text = usermapping_user_sAMAccountName.SelectedItem.ToString();
                bob.SubItems["columnIndex"].Tag = usermapping_user_sAMAccountName.SelectedIndex;
                bob.SubItems["whereText"].Text = usermapping_user_where.Text;
                bob.SubItems["whereText"].Tag = usermapping_user_where.Text;
            }
            else
            {
                //Add
                if (usermapping_user_table_view.Text == "")
                    usermapping_user_table_view.Focus();
                else if (usermapping_user_source.Text == "")
                    usermapping_user_source.Focus();
                else if (usermapping_user_sAMAccountName.Text == "")
                    usermapping_user_sAMAccountName.Focus();
                else
                {
                    ListViewItem bob = new ListViewItem();
                    ListViewItem.ListViewSubItem jane = new ListViewItem.ListViewSubItem();
                    bob.Text = userMapping_factName.Text;
                    bob.Name = userMapping_factName.Text;
                    jane.Name = "serverName";
                    jane.Text = userMapping_DBServerName.Text;
                    jane.Tag = userMapping_DBServerName.Text;
                    bob.SubItems.Add(jane);
                    jane = new ListViewItem.ListViewSubItem();
                    jane.Name = "databaseIndex";
                    jane.Text = userMapping_DatabaseName.SelectedItem.ToString();
                    jane.Tag = userMapping_DatabaseName.SelectedIndex;
                    bob.SubItems.Add(jane);
                    jane = new ListViewItem.ListViewSubItem();
                    jane.Name = "typeIndex";
                    jane.Text = usermapping_user_table_view.SelectedItem.ToString();
                    jane.Tag = usermapping_user_table_view.SelectedIndex;
                    bob.SubItems.Add(jane);
                    jane = new ListViewItem.ListViewSubItem();
                    jane.Name = "tableIndex";
                    jane.Text = usermapping_user_source.SelectedItem.ToString();
                    jane.Tag = usermapping_user_source.SelectedIndex;
                    bob.SubItems.Add(jane);
                    jane = new ListViewItem.ListViewSubItem();
                    jane.Name = "columnIndex";
                    jane.Text = usermapping_user_sAMAccountName.SelectedItem.ToString();
                    jane.Tag = usermapping_user_sAMAccountName.SelectedIndex; ;
                    bob.SubItems.Add(jane);
                    jane = new ListViewItem.ListViewSubItem();
                    jane.Name = "whereText";
                    jane.Text = usermapping_user_where.Text;
                    jane.Tag = usermapping_user_where.Text;
                    bob.SubItems.Add(jane);

                    if (bob.Text.Trim().Length == 0 || userMapping_factList.Items.ContainsKey(bob.Text))
                    {
                        MessageBox.Show("I'm Sorry but every fact is required to have a name that is unique.");
                    }
                    else
                    {
                        userMapping_factList.Items.Add(bob);
                        userMapping_fact_Add_Edit.Text = "Add";
                        userMapping_factName.Text = "";
                        userMapping_DBServerName.Enabled = false;
                        userMapping_DatabaseName.Enabled = false;
                        usermapping_user_table_view.SelectedIndex = -1;
                        usermapping_user_source.DataSource = null;
                        usermapping_user_sAMAccountName.DataSource = null;
                        usermapping_user_where.Text = "";
                    }
                }
            }
        }
        private void userMapping_factList_ClientSizeChanged(object sender, EventArgs e)
        {
            userMapping_factHeading.Width = userMapping_factList.ClientSize.Width;
        }
        private void userMapping_factList_SelectedIndexChanged(object sender, EventArgs e)
        {
            userMapping_UserData.Enabled = true;
            userMapping_fact_Add_Edit.Enabled = true;
            userMapping_fact_Add_Edit.Text = "Add";
            userMapping_factName.Text = "";
            usermapping_user_table_view.SelectedIndex = -1;
            usermapping_user_source.DataSource = null;
            usermapping_user_sAMAccountName.DataSource = null;
            usermapping_user_where.Text = "";

            if (userMapping_factList.SelectedItems.Count == 1)
            {
                userMapping_factDelete.Enabled = true;
                //This Doesn't Work right and I don't know why.

                userMapping_fact_Add_Edit.Text = "Edit/Save";
                ListViewItem bob = userMapping_factList.SelectedItems[0];
                userMapping_factName.Text                     = bob.Text;
                usermapping_user_table_view.SelectedIndex = -1;
                usermapping_user_table_view.SelectedIndex     = Convert.ToInt32(bob.SubItems["typeIndex"].Tag);
                usermapping_user_source.SelectedIndex         = Convert.ToInt32(bob.SubItems["tableIndex"].Tag);
                usermapping_user_sAMAccountName.SelectedIndex = Convert.ToInt32(bob.SubItems["columnIndex"].Tag);
                usermapping_user_where.Text = Convert.ToString(bob.SubItems["whereText"].Tag);
            }
            else if (userMapping_factList.SelectedItems.Count > 1)
            {
                userMapping_UserData.Enabled = false;
                userMapping_fact_Add_Edit.Enabled = false;
            }

            else
            {
                userMapping_factDelete.Enabled = false;
            }
        }
        private void userMapping_DBServerName_Leave(object sender, EventArgs e)
        {
            if (userMapping_DBServerName.Text != "")
            {
                try
                {
                    ArrayList tableList = new ArrayList();
                    System.Data.SqlClient.SqlConnection SqlCon = new System.Data.SqlClient.SqlConnection("server=" + userMapping_DBServerName.Text + ";Integrated Security=SSPI;");
                    SqlCon.Open();
                    System.Data.SqlClient.SqlCommand SqlCom = new System.Data.SqlClient.SqlCommand();
                    SqlCom.Connection = SqlCon;
                    SqlCom.CommandType = CommandType.StoredProcedure;
                    SqlCom.CommandText = "sp_databases";

                    System.Data.SqlClient.SqlDataReader r;
                    r = SqlCom.ExecuteReader();

                    while (r.Read())
                    {
                        tableList.Add((string)r[0].ToString().Trim());
                    }
                    r.Close();
                    SqlCon.Close();

                    userMapping_DatabaseName.DataSource = tableList;
                }
                catch
                {
                    MessageBox.Show("Error Connecting to Database Server.  Perhaps the server name was mis entered?");
                    userMapping_DBServerName.Text = "";
                }
        }


        }
        private void userMapping_DatabaseName_Enter(object sender, EventArgs e)
        {
            if (userMapping_DBServerName.Text == "")
                userMapping_DBServerName.Focus();
        }
        private void userMapping_DBServerName_Enter(object sender, EventArgs e)
        {
            if (userMapping_factName.Text == "")
                userMapping_factName.Focus();
        }
        private void usermapping_user_table_view_Enter(object sender, EventArgs e)
        {
            if (userMapping_DatabaseName.Text == "")
                userMapping_DatabaseName.Focus();
        }
        private void usermapping_user_source_Enter(object sender, EventArgs e)
        {
            if (usermapping_user_table_view.Text == "")
                usermapping_user_table_view.Focus();
        }
        private void usermapping_user_sAMAccountName_Enter(object sender, EventArgs e)
        {
            if (usermapping_user_source.Text == "")
                usermapping_user_source.Focus();
        }
        private void usermapping_user_where_Enter(object sender, EventArgs e)
        {
            if (usermapping_user_sAMAccountName.Text == "")
                usermapping_user_sAMAccountName.Focus();
        }
        private void userMapping_DBServerName_TextChanged(object sender, EventArgs e)
        {
            if (userMapping_DBServerName.Focused == false)
                userMapping_DBServerName_Leave(null, null);
        }
        private void mappinggrid_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
                // added to ignore the data error when working with an unbound datagridview and setting the datagridviewcombobox to have a selected value
        }
        private void userMapping_factDelete_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem deletable in userMapping_factList.SelectedItems) {
                deletable.Remove();
            }
            if (userMapping_factList.Items.Count == 0)
            {
                userMapping_DBServerName.Enabled = true;
                userMapping_DatabaseName.Enabled = true;
            }
        }
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1 && userMapping_factList.Items.Count == 0)
            {
                tabControl1.SelectedIndex = 0;
                MessageBox.Show("You must complete Step 1 before starting Step 2.");
            }
            if (userMapping_factList.Items.Count > 0)
            {
                userMapping_factSelectBox.Items.Clear();
                foreach (ListViewItem temp in userMapping_factList.Items)
                {
                    userMapping_factSelectBox.Items.Add(temp.Text);
                }
            }
        }

		//UI DIALOG  DATA ENTRY FOR GMAIL TAB

		private void Domain_TextChanged(object sender, EventArgs e)
        {
            label48.Text = "User@" + mailDomain.Text.ToString();
			guserconfig.Admin_domain = mailDomain.Text.ToString();
			
        }
        private void mailUser_TextChanged(object sender, EventArgs e)
        {
            label48.Text = mailUser.Text.ToString() + "@" + mailDomain.Text.ToString();
			guserconfig.Admin_user = mailUser.Text.ToString();
        }   
        private void mailPassword_TextChanged(object sender, EventArgs e)
        {
            if (mailPassword.Text != "")
            {
                label49.Text = "Press Below To Check Authentication";
                mailCheckAuth.Visible = true;
				guserconfig.Admin_password = mailPassword.Text.ToString();
            }
            else
            {
                label49.Text = "Complete Login Info To Authenticate";
                mailCheckAuth.Visible = false;
            }
        } 
        private void mailUser_Enter(object sender, EventArgs e)
        {
            if (mailDomain.Text.ToString() == "")
            {
                mailDomain.Focus();
            }
        }
        private void mailPassword_Enter(object sender, EventArgs e)
        {
            if (mailUser.Text.ToString() == "")
            {
                mailUser.Focus();
            }
        }

        private void mail_user_AD_or_SQL_SelectedIndexChanged(object sender, EventArgs e)
        {
            guserconfig.User_Datasource = mail_user_AD_or_SQL.Text.ToString();
            ArrayList adobjects = new ArrayList();
            if (mail_user_AD_or_SQL.Text.ToString() == "Active Directory")
            {
                mail_fields_Fname.DataSource = null;
                mail_fields_password.DataSource = null;
                mail_fields_userID.DataSource = null;
                mail_fields_Mname.DataSource = null;
                mail_fields_Lname.DataSource = null;

                label50.Visible = false;
                label51.Visible = false;
                mail_user_source.Visible = false;
                mail_user_Table_View.Visible = false;
                mail_user_where.Visible = false;
                groupBox13.Text = "Active Directory OU Info";
                mail_user_OU.Visible = true;
                label149.Visible = true;
                adobjects = tools.ADobjectAttribute();
                mail_fields_Fname.DataSource = adobjects.Clone();
                mail_fields_password.DataSource = adobjects.Clone();
                mail_fields_userID.DataSource = adobjects.Clone();
                mail_fields_Mname.DataSource = adobjects.Clone();
                mail_fields_Lname.DataSource = adobjects.Clone();


            }
            else if (mail_user_AD_or_SQL.Text.ToString() == "Database")
            {
                mail_fields_Fname.DataSource = null;
                mail_fields_password.DataSource = null;
                mail_fields_userID.DataSource = null;
                mail_fields_Mname.DataSource = null;
                mail_fields_Lname.DataSource = null;

                label50.Visible = true;
                label51.Visible = true;
                mail_user_source.Visible = true;
                mail_user_Table_View.Visible = true;
                mail_user_where.Visible = true;
                groupBox13.Text = "Database Table Info";
                mail_user_OU.Visible = false;
                label149.Visible = false; 
            }
            
        }
        private void mail_user_Table_View_SelectedIndexChanged(object sender, EventArgs e)
        {
            // only valid for SQL server 2000
            if (guserconfig.DBCatalog != "" && guserconfig.DataServer != "")
            {
                //populates table dialog with tables or views depending on the results of a query
                ArrayList tableList = new ArrayList();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + guserconfig.DataServer.ToString() + ";Initial Catalog=" + guserconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");
                SqlCommand sqlComm;
                sqlConn.Open();
                // create the command object
                if (mail_user_Table_View.Text.ToLower() == "table")
                {
                    sqlComm = new SqlCommand("SELECT name FROM sysobjects where TYPE = 'U' order by NAME", sqlConn);
                }
                else
                {
                    sqlComm = new SqlCommand("SELECT name from SYSOBJECTS where TYPE = 'V' order by NAME", sqlConn);
                }
                SqlDataReader r = sqlComm.ExecuteReader();
                while (r.Read())
                {
                    tableList.Add((string)r[0].ToString().Trim());
                }
                r.Close();
                sqlConn.Close();

                guserconfig.User_table_view = mail_user_Table_View.Text.ToString();
                mail_user_source.DataSource = tableList;
            }
            else
            {
                MessageBox.Show("Please set the dataserver and catalog");
            }
        }
        private void mail_user_source_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (guserconfig.DBCatalog != "" && guserconfig.DataServer != "")
            {
                //populates columns dialog with columns depending on the results of a query
                ArrayList columnList = new ArrayList();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + guserconfig.DataServer.ToString() + ";Initial Catalog=" + guserconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");

                sqlConn.Open();
                // create the command object
                SqlCommand sqlComm = new SqlCommand("SELECT column_name FROM information_schema.columns WHERE table_name = '" + mail_user_source.Text.ToString() + "'", sqlConn);
                SqlDataReader r = sqlComm.ExecuteReader();
                while (r.Read())
                {
                    columnList.Add((string)r[0].ToString().Trim());
                }
                r.Close();
                sqlConn.Close();

                guserconfig.User_dbTable = mail_user_source.Text.ToString();
                mail_fields_Fname.DataSource = columnList;
                mail_fields_password.DataSource = columnList.Clone();
                mail_fields_userID.DataSource = columnList.Clone();
                mail_fields_Mname.DataSource = columnList.Clone();
                mail_fields_Lname.DataSource = columnList.Clone();
            }
            else
            {
                MessageBox.Show("Please select table or view");
            }
        }
        private void mail_user_where_TextChanged(object sender, EventArgs e)
        {
            guserconfig.User_where = mail_user_where.Text.ToString();
        }         
        private void mail_user_OU_TextChanged(object sender, EventArgs e)
        {
            guserconfig.User_ad_OU = mail_user_OU.Text.ToString();
        }

		private void mail_fields_userID_SelectedIndexChanged(object sender, EventArgs e)
		{
			guserconfig.User_StuID = mail_fields_userID.Text.ToString();
		}
		private void mail_fields_password_SelectedIndexChanged(object sender, EventArgs e)
		{
			guserconfig.User_password = mail_fields_password.Text.ToString();
		}
		private void mail_fields_Fname_SelectedIndexChanged(object sender, EventArgs e)
		{
			guserconfig.User_Fname = mail_fields_Fname.Text.ToString();
		}
		private void mail_fields_Lname_SelectedIndexChanged(object sender, EventArgs e)
		{
			guserconfig.User_Lname = mail_fields_Lname.Text.ToString();
		}
        private void mail_fields_Mname_SelectedIndexChanged(object sender, EventArgs e)
        {
            guserconfig.User_Mname = mail_fields_Mname.Text.ToString();
        }

        private void mail_writeback_Database_CheckedChanged(object sender, EventArgs e)
        {
            if (mail_writeback_database.Checked == true)
            {
                label150.Enabled = true;
                label147.Enabled = true;
                label151.Enabled = true;
                label153.Enabled = true;
                label154.Enabled = true;
                label155.Enabled = true;
                label156.Enabled = true;
                mail_writeback_table.Enabled = true;
                mail_writeback_key.Enabled = true;
                mail_writeback_use_secondary_email.Enabled = true;
                mail_writeback_secondary_email.Enabled = true;
                mail_writeback_where.Enabled = true;
                mail_writeback_email.Enabled = true;
                
				// only valid for SQL server 2000
				if (guserconfig.DBCatalog != "" && guserconfig.DataServer != "" && mail_writeback_table.DataSource == null)
				{
					//populates table dialog with tables or views depending on the results of a query
					ArrayList tableList = new ArrayList();
					SqlConnection sqlConn = new SqlConnection("Data Source=" + guserconfig.DataServer.ToString() + ";Initial Catalog=" + guserconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");
					SqlCommand sqlComm;
					sqlConn.Open();
					// create the command object
					sqlComm = new SqlCommand("SELECT name FROM sysobjects where TYPE = 'U' order by NAME", sqlConn);
					SqlDataReader r = sqlComm.ExecuteReader();
					while (r.Read())
					{
						tableList.Add((string)r[0].ToString().Trim());
					}
					r.Close();
					sqlConn.Close();

					guserconfig.Writeback_DB_checkbox = true;
					mail_writeback_table.DataSource = tableList;
				}
				else if (guserconfig.DBCatalog == "" && guserconfig.DataServer == "")
				{
					MessageBox.Show("Please set the dataserver and catalog");
				}

            }
            if (mail_writeback_database.Checked == false)
            {
                label150.Enabled = false;
                label147.Enabled = false;
                label151.Enabled = false;
                label153.Enabled = false;
                label154.Enabled = false;
                label155.Enabled = false;
                label156.Enabled = false;
                mail_writeback_table.Enabled = false;
                mail_writeback_key.Enabled = false;
                mail_writeback_use_secondary_email.Enabled = false;
                mail_writeback_secondary_email.Enabled = false;
                mail_writeback_where.Enabled = false;
                mail_writeback_email.Enabled = false;
				guserconfig.Writeback_DB_checkbox = false;

            }
        } 
        private void mail_writeback_Active_directory_CheckedChanged(object sender, EventArgs e)
        {
            if (mail_writeback_Active_directory.Checked == true)
            {
                label152.Enabled = true;
                mail_writeback_ou.Enabled = true;
				guserconfig.Writeback_AD_checkbox = true;
            }
            if (mail_writeback_Active_directory.Checked == false)
            {
                label152.Enabled = false;
                mail_writeback_ou.Enabled = false;
				guserconfig.Writeback_AD_checkbox = false;
            }
        }
        private void mail_writeback_table_SelectedIndexChanged(object sender, EventArgs e)
        {
			if (guserconfig.DBCatalog != "" && guserconfig.DataServer != "")
			{
				//populates columns dialog with columns depending on the results of a query
				ArrayList columnList = new ArrayList();
				SqlConnection sqlConn = new SqlConnection("Data Source=" + guserconfig.DataServer.ToString() + ";Initial Catalog=" + guserconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");

				sqlConn.Open();
				// create the command object
				SqlCommand sqlComm = new SqlCommand("SELECT column_name FROM information_schema.columns WHERE table_name = '" + mail_writeback_table.Text.ToString() + "'", sqlConn);
				SqlDataReader r = sqlComm.ExecuteReader();
				while (r.Read())
				{
					columnList.Add((string)r[0].ToString().Trim());
				}
				r.Close();
				sqlConn.Close();

				guserconfig.Writeback_table = mail_writeback_table.Text.ToString();
				mail_writeback_email.DataSource = columnList;
                mail_writeback_secondary_email.DataSource = columnList.Clone();
                mail_writeback_key.DataSource = columnList.Clone();   
			}
        }
        private void mail_writeback_key_SelectedIndexChanged(object sender, EventArgs e)
        {
            guserconfig.Writeback_primary_key = mail_writeback_key.Text.ToString();
        }
        private void mail_writeback_use_secondary_email_CheckedChanged(object sender, EventArgs e)
        {
            if (mail_writeback_use_secondary_email.Checked == true)
            {
                mail_writeback_secondary_email.Visible = true;
                label154.Visible = true;
                label155.Visible = false;
                label156.Visible = false;
                guserconfig.Writeback_transfer_email_checkbox = true;
            }
            else
            {
                mail_writeback_secondary_email.Visible = false;
                label154.Visible = false;
                label155.Visible = true;
                label156.Visible = true;
                guserconfig.Writeback_transfer_email_checkbox = false;
            }
        }
        private void mail_writeback_where_TextChanged(object sender, EventArgs e)
        {
            guserconfig.Writeback_where_clause = mail_writeback_where.Text.ToString();
        }
        private void mail_writeback_email_SelectedIndexChanged(object sender, EventArgs e)
        {
			guserconfig.Writeback_email_field = mail_writeback_email.Text.ToString();
        }
        private void mail_writeback_secondary_email_SelectedIndexChanged(object sender, EventArgs e)
        {
            guserconfig.Writeback_secondary_email_field = mail_writeback_secondary_email.Text.ToString();
        }
        private void mail_writeback_ou_TextChanged(object sender, EventArgs e)
        {
			guserconfig.Writeback_ad_OU = mail_writeback_ou.Text.ToString();
        }                          


        // BUTTONS FOR GMAIL TAB
        private void gmail_Check_Authorization_Click(object sender, EventArgs e)
        {
            authNotify.Text = "Contacting google";
            AppsService service = new AppsService(mailDomain.Text.ToString(), mailUser.Text.ToString() + "@" + mailDomain.Text.ToString(), mailPassword.Text.ToString());

            try
            {
                authNotify.Text = "Checking for user";
                UserEntry gAdmin = service.RetrieveUser(mailUser.Text.ToString());
                authNotify.Text = "Successfully Authenticated! \n" + gAdmin.Login.UserName.ToString();

            }
            catch
            {                
                authNotify.Text = "Failed Authentication";
            }
        }
		private void mail_save_Click(object sender, EventArgs e)
		{
			Dictionary<string, string> properties = new Dictionary<string, string>();
			int i = 0;
			SaveFileDialog saveFileDialog1 = new SaveFileDialog();
			saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
			saveFileDialog1.FilterIndex = 2;
			saveFileDialog1.RestoreDirectory = true;
			if (saveFileDialog1.ShowDialog() == DialogResult.OK)
			{
				// create a file stream, where "c:\\testing.txt" is the file path
				System.IO.FileStream fs = new System.IO.FileStream(saveFileDialog1.FileName, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write, System.IO.FileShare.ReadWrite);

				// create a stream writer
				System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.ASCII);

				// write to file (buffer), where textbox1 is your text box
				properties = guserconfig.ToDictionary();
				ICollection<string> c = properties.Keys;
				i = c.Count;
				foreach (string str in c)
				{
					sw.WriteLine("{0} | {1:C}", str, properties[str]);
				}

				// flush buffer (so the text really goes into the file)
				sw.Flush();

				// close stream writer and file
				sw.Close();
				fs.Close();
			}
		}
		private void mail_open_Click(object sender, EventArgs e)
		{
			Dictionary<string, string> properties = new Dictionary<string, string>();
			BindingSource bs = new BindingSource();

			OpenFileDialog openFileDialog1 = new OpenFileDialog();
			openFileDialog1.InitialDirectory = "c:\\";
			openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
			openFileDialog1.FilterIndex = 2;
			openFileDialog1.RestoreDirectory = true;
			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				StreamReader re = File.OpenText(openFileDialog1.FileName);

				string input = null;
				while ((input = re.ReadLine()) != null)
				{
					string[] parts = input.Split('|');
					properties.Add(parts[0].Trim(), parts[1].Trim());
				}
				re.Close();
			}

			// Load values into text boxes
			// reload properties each time as they are overwritten with the combo object trigger events
			guserconfig.load(properties);
            DBserver.Text = guserconfig.DataServer;
            guserconfig.load(properties);
            Catalog.Text = guserconfig.DBCatalog;
            guserconfig.load(properties);
			//mail = userconfig.DataServer;
            mailDomain.Text = guserconfig.Admin_domain;
			guserconfig.load(properties);
            mailUser.Text = guserconfig.Admin_user;
			guserconfig.load(properties);
            mailPassword.Text = guserconfig.Admin_password;
            guserconfig.load(properties);
            
            mail_user_AD_or_SQL.Text = guserconfig.User_Datasource;
            guserconfig.load(properties);
            mail_user_Table_View.Text = guserconfig.User_table_view;
            guserconfig.load(properties);
            mail_user_source.Text = guserconfig.User_dbTable;
            guserconfig.load(properties);
            mail_user_where.Text = guserconfig.User_where;
            guserconfig.load(properties);
            mail_user_OU.Text = guserconfig.User_ad_OU;
            guserconfig.load(properties);

            mail_fields_Fname.Text = guserconfig.User_Fname;
            guserconfig.load(properties);
            mail_fields_Lname.Text = guserconfig.User_Lname;
            guserconfig.load(properties);
            mail_fields_Mname.Text = guserconfig.User_Mname;
            guserconfig.load(properties);
            mail_fields_password.Text = guserconfig.User_password;
            guserconfig.load(properties);
            mail_fields_userID.Text = guserconfig.User_StuID;
            guserconfig.load(properties);

            mail_writeback_Active_directory.Checked = guserconfig.Writeback_AD_checkbox;
            guserconfig.load(properties);
            mail_writeback_database.Checked = guserconfig.Writeback_DB_checkbox;
            guserconfig.load(properties);
            mail_writeback_table.Text = guserconfig.Writeback_table;
            guserconfig.load(properties);
            mail_writeback_key.Text = guserconfig.Writeback_primary_key;
            guserconfig.load(properties);
            mail_writeback_use_secondary_email.Checked = guserconfig.Writeback_transfer_email_checkbox;
            guserconfig.load(properties);
            mail_writeback_where.Text = guserconfig.Writeback_where_clause;
            guserconfig.load(properties);
            mail_writeback_email.Text = guserconfig.Writeback_email_field;
            guserconfig.load(properties);
            mail_writeback_secondary_email.Text = guserconfig.Writeback_secondary_email_field;
            guserconfig.load(properties);
            mail_writeback_ou.Text = guserconfig.Writeback_ad_OU;
            guserconfig.load(properties);
			
		}
		private void mail_execute_Click(object sender, EventArgs e)
		{
			gmailSyncr.EmailUsersSync(guserconfig, tools, log);
		}



        







        // custom AD fields 
    }
}
