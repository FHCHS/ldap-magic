using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using WindowsApplication1;
using WindowsApplication1.utils;

namespace WindowsApplication1
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Arguments CommandLine = new Arguments(args);
            string operation = "";
            string file = "";
            if (CommandLine["F"] != null)
            {
                file = CommandLine["F"];
            }
            else Console.WriteLine("File not defined full path needed -F=c:\\xxx\\yyy\\zzz\\file.sav warning there is no sanity check on the path");

            operation = CommandLine["O"];
            if (CommandLine["O"] != null)
            {
                if (CommandLine["O"] != "users" || CommandLine["O"] != "groups" || CommandLine["O"] != "OUmap" || CommandLine["O"] != "gmail")
                {
                    Console.WriteLine("Operation not defined -O=users -O=groups -O=OUmap -O=gmail");
                }
                else
                {
                    operation = CommandLine["O"];
                }
            }
            // MessageBox.Show("operation is " + operation + " file is " + file);
            if (file != "" && operation != "")
            {

                // woot halleluijah we have input from the user time to execute
                //duplicate the gui fucntionality in cmd line 
                // we won't check this input cause its from a really smart system administrator
                // just in case file expects a full path 
                // c:\blah\blah\blah.ext
                // valid oprations are 
                // users	groups	 OUmap	 gmail
                // create objects to hold save data

                GroupSynch groupconfig = new GroupSynch();
                UserSynch userconfig = new UserSynch();
                GmailUsers guserconfig = new GmailUsers();
                executionOrder execution = new executionOrder();
                UserStateChange usermapping = new UserStateChange();
                utils.ToolSet tools = new ToolSet();
                LogFile log = new LogFile();
                ObjectADSqlsyncGroup groupSyncr = new ObjectADSqlsyncGroup();
                ObjectADGoogleSync gmailSyncr = new ObjectADGoogleSync();

                // perform operations based on the data input from the user fro groups users, OU's and gmail
                if (operation == "group")
                {
                    Dictionary<string, string> properties = new Dictionary<string, string>();

                    StreamReader re = File.OpenText(file);
                    string input = null;
                    while ((input = re.ReadLine()) != null)
                    {
                        string[] parts = input.Split('|');
                        properties.Add(parts[0].Trim(), parts[1].Trim());
                    }
                    re.Close();
                    // Load values into text boxes
                    // reload properties each time as they are overwritten with the combo object trigger events
                    groupconfig.Load(properties);
                    groupSyncr.ExecuteGroupSync(groupconfig, tools, log);


                    //// save log to disk
                    //SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    //saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                    //saveFileDialog1.FilterIndex = 2;
                    //saveFileDialog1.RestoreDirectory = true;
                    //if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    //{
                    //    // create a file stream, where "c:\\testing.txt" is the file path
                    //    System.IO.FileStream fs = new System.IO.FileStream(saveFileDialog1.FileName, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write, System.IO.FileShare.ReadWrite);

                    //    // create a stream writer
                    //    System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.ASCII);

                    //    // write to file (buffer), where textbox1 is your text box
                    //    sw.WriteLine("{0}", result2);
                    //    sw.WriteLine("{0}", result);


                    //    // flush buffer (so the text really goes into the file)
                    //    sw.Flush();

                    //    // close stream writer and file
                    //    sw.Close();
                    //    fs.Close();
                    //}

                }
                if (operation == "users")
                {
                    Dictionary<string, string> properties = new Dictionary<string, string>();
                    DataTable customs = new DataTable();
                    BindingSource bs = new BindingSource();

                    //OpenFileDialog openFileDialog1 = new OpenFileDialog();
                    //openFileDialog1.InitialDirectory = "c:\\";
                    //openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                    //openFileDialog1.FilterIndex = 2;
                    //openFileDialog1.RestoreDirectory = true;
                    //if (openFileDialog1.ShowDialog() == DialogResult.OK)
                    //{
                    try
                    {
                        StreamReader re = File.OpenText(file);

                        string input = null;
                        while ((input = re.ReadLine()) != null)
                        {
                            string[] parts = input.Split('|');
                            properties.Add(parts[0].Trim(), parts[1].Trim());
                        }
                        re.Close();
                    }
                    catch
                    {
                        Console.Write("failed file load");
                    }
                    //}

                    // Load values into text boxes
                    // reload properties each time as they are overwritten with the combo object trigger events
                    userconfig.load(properties);
                    //MessageBox.Show("executing");
                    groupSyncr.ExecuteUserSync(userconfig, tools, log);
                    //MessageBox.Show("complete");
                    //StringBuilder result = new StringBuilder();
                    //int i = 0;
                    //result.Append("***************************\n*                         *\n*        Transactions     *\n*                         *\n***************************");
                    //for (i = 0; i < log.transactions.Count; i++)
                    //{
                    //    result.Append(log.transactions[i].ToString() + "\n");
                    //}
                    //result.Append("***************************\n*                         *\n*        Warnings         *\n*                         *\n***************************");

                    //for (i = 0; i < log.warnings.Count; i++)
                    //{
                    //    result.Append(log.warnings[i].ToString() + "\n");
                    //}
                    //result.Append("***************************\n*                         *\n*        Errors           *\n*                         *\n***************************");
                    //for (i = 0; i < log.errors.Count; i++)
                    //{
                    //    result.Append(log.errors[i].ToString() + "\n");
                    //}
                    //MessageBox.Show(result.ToString());

                    //// save log to disk
                    //SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    //saveFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                    //saveFileDialog1.FilterIndex = 2;
                    //saveFileDialog1.RestoreDirectory = true;
                    //if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    //{
                    //    // create a file stream, where "c:\\testing.txt" is the file path
                    //    System.IO.FileStream fs = new System.IO.FileStream(saveFileDialog1.FileName, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write, System.IO.FileShare.ReadWrite);

                    //    // create a stream writer
                    //    System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.ASCII);

                    //    // write to file (buffer), where textbox1 is your text box
                    //    sw.WriteLine("{0}", result2);
                    //    sw.WriteLine("{0}", result);


                    //    // flush buffer (so the text really goes into the file)
                    //    sw.Flush();

                    //    // close stream writer and file
                    //    sw.Close();
                    //    fs.Close();
                    //}
                }
                if (operation == "gmail")
                {
                    Dictionary<string, string> properties = new Dictionary<string, string>();
                    BindingSource bs = new BindingSource();
                    try
                    {
                        StreamReader re = File.OpenText(file);

                        string input = null;
                        while ((input = re.ReadLine()) != null)
                        {
                            string[] parts = input.Split('|');
                            properties.Add(parts[0].Trim(), parts[1].Trim());
                        }
                        re.Close();
                    }
                    catch
                    {
                        Console.Write("failed file load");
                    }
                    guserconfig.load(properties);
                    gmailSyncr.EmailUsersSync(guserconfig, tools, log);
                }

            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
        }
    }
}