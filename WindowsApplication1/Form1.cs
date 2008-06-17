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
        public enum objectClass
        {
            user, group, computer
        }
        public enum returnType
        {
            distinguishedName, ObjectGUID
        }
        public enum accountFlags
        {
            ADS_UF_SCRIPT = 0x1,
            ADS_UF_ACCOUNTDISABLE = 0x2,
            ADS_UF_HOMEDIR_REQUIRED = 0x8,
            ADS_UF_LOCKOUT = 0x10,
            ADS_UF_PASSWD_NOTREQD = 0x20,
            ADS_UF_PASSWD_CANT_CHANGE = 0x40,
            ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = 0x80,
            ADS_UF_TEMP_DUPLICATE_ACCOUNT = 0x100,
            ADS_UF_NORMAL_ACCOUNT = 0x200,
            ADS_UF_INTERDOMAIN_TRUST_ACCOUNT = 0x800,
            ADS_UF_WORKSTATION_TRUST_ACCOUNT = 0x1000,
            ADS_UF_SERVER_TRUST_ACCOUNT = 0x2000,
            ADS_UF_DONT_EXPIRE_PASSWD = 0x10000,
            ADS_UF_MNS_LOGON_ACCOUNT = 0x20000,
            ADS_UF_SMARTCARD_REQUIRED = 0x40000,
            ADS_UF_TRUSTED_FOR_DELEGATION = 0x80000,
            ADS_UF_NOT_DELEGATED = 0x100000,
            ADS_UF_USE_DES_KEY_ONLY = 0x200000,
            ADS_UF_DONT_REQUIRE_PREAUTH = 0x400000,
            ADS_UF_PASSWORD_EXPIRED = 0x800000,
            ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = 0x1000000
        }
        public Form1()
        {
            InitializeComponent();
        }
        public class LogFile
        {
            private List<string> logtransactions;
            private List<string> logterrors;
            private List<string> logwarnings;
            public LogFile()
            {
                logtransactions = new List<string>();
                logterrors = new List<string>();
                logwarnings = new List<string>();
            }

            public List<string> transactions
            {
                get
                {
                    return logtransactions;
                }
                set
                {
                    logtransactions = value;
                }
            }
            public List<string> errors
            {
                get
                {
                    return logterrors;
                }
                set
                {
                    logterrors = value;
                }
            }
            public List<string> warnings
            {
                get
                {
                    return logwarnings;
                }
                set
                {
                    logwarnings = value;
                }
            }


        }
        public class GroupSynch
        {

            private String configBaseGroupOU;
            private String configBaseUserOU;
            private String configNotes;
            private String configGroup_CN;
            private String configGroup_table_view;
            private String configGroup_sAMAccount;
            private String configGroup_dbTable;
            private String configGroup_where;
            private String configGroup_Prepend;
            private String configUser_Group_Reference;
            private String configUser_table_view;
            private String configUser_sAMAccount;
            private String configUser_dbTable;
            private String configUser_where;
            private String configDataServer;
            private String configDBCatalog;
            private String configprogress;

            // constructor creates blank strings
            public GroupSynch()
            {
                configBaseGroupOU = "";
                configBaseUserOU = "";
                configNotes = "";
                configGroup_CN = "";
                configGroup_table_view = "";
                configGroup_sAMAccount = "";
                configGroup_dbTable = "";
                configGroup_where = "";
                configUser_Group_Reference = "";
                configUser_table_view = "";
                configUser_sAMAccount = "";
                configUser_dbTable = "";
                configUser_where = "";
                configDataServer = "";
                configDBCatalog = "";
            }
            public GroupSynch(Dictionary<string, string> dictionary)
            {
                configBaseGroupOU = dictionary[configBaseGroupOU].ToString();
                configBaseUserOU = dictionary[configBaseUserOU].ToString();
                configNotes = dictionary[configNotes].ToString();
                configGroup_CN = dictionary[configGroup_CN].ToString();
                configGroup_table_view = dictionary[configGroup_table_view].ToString();
                configGroup_sAMAccount = dictionary[configGroup_sAMAccount].ToString();
                configGroup_dbTable = dictionary[configGroup_dbTable].ToString();
                configGroup_where = dictionary[configGroup_where].ToString();
                configUser_Group_Reference = dictionary[configUser_Group_Reference].ToString();
                configUser_table_view = dictionary[configUser_table_view].ToString();
                configUser_sAMAccount = dictionary[configUser_sAMAccount].ToString();
                configUser_dbTable = dictionary[configUser_dbTable].ToString();
                configUser_where = dictionary[configUser_where].ToString();
                configDataServer = dictionary[configDataServer].ToString();
                configDBCatalog = dictionary[configDBCatalog].ToString();
            }

            public void Load(Dictionary<string, string> dictionary)
            {
                dictionary.TryGetValue("configBaseGroupOU", out configBaseGroupOU);
                dictionary.TryGetValue("configBaseUserOU", out configBaseUserOU);
                dictionary.TryGetValue("configNotes", out configNotes);
                dictionary.TryGetValue("configGroup_CN", out configGroup_CN);
                dictionary.TryGetValue("configGroup_table_view", out configGroup_table_view);
                dictionary.TryGetValue("configGroup_sAMAccount", out configGroup_sAMAccount);
                dictionary.TryGetValue("configGroup_dbTable", out configGroup_dbTable);
                dictionary.TryGetValue("configGroup_where", out configGroup_where);
                dictionary.TryGetValue("configGroup_Prepend", out configGroup_Prepend);
                dictionary.TryGetValue("configUser_Group_Reference", out configUser_Group_Reference);
                dictionary.TryGetValue("configUser_table_view", out configUser_table_view);
                dictionary.TryGetValue("configUser_sAMAccount", out configUser_sAMAccount);
                dictionary.TryGetValue("configUser_dbTable", out configUser_dbTable);
                dictionary.TryGetValue("configUser_where", out configUser_where);
                dictionary.TryGetValue("configDataServer", out configDataServer);
                dictionary.TryGetValue("configDBCatalog", out configDBCatalog);
            }

            // accessor to properties
            public String progress
            {
                get
                {
                    return configprogress;
                }
                set
                {
                    configprogress = value;
                }
            }
            public String BaseGroupOU
            {
                get
                {
                    return configBaseGroupOU;
                }
                set
                {
                    configBaseGroupOU = value;
                }
            }
            public String BaseUserOU
            {
                get
                {
                    return configBaseUserOU;
                }
                set
                {
                    configBaseUserOU = value;
                }
            }
            public String Notes
            {
                get
                {
                    return configNotes;
                }
                set
                {
                    configNotes = value;
                }
            }
            public String Group_CN
            {
                get
                {
                    return configGroup_CN;
                }
                set
                {
                    configGroup_CN = value;
                }
            }
            public String Group_table_view
            {
                get
                {
                    return configGroup_table_view;
                }
                set
                {
                    configGroup_table_view = value;
                }
            }
            public String Group_sAMAccount
            {
                get
                {
                    return configGroup_sAMAccount;
                }
                set
                {
                    configGroup_sAMAccount = value;
                }
            }
            public String Group_dbTable
            {
                get
                {
                    return configGroup_dbTable;
                }
                set
                {
                    configGroup_dbTable = value;
                }
            }
            public String Group_where
            {
                get
                {
                    return configGroup_where;
                }
                set
                {
                    configGroup_where = value;
                }
            }
            public String Group_Append
            {
                get
                {
                    return configGroup_Prepend;
                }
                set
                {
                    configGroup_Prepend = value;
                }
            }
            public String User_Group_Reference
            {
                get
                {
                    return configUser_Group_Reference;
                }
                set
                {
                    configUser_Group_Reference = value;
                }
            }
            public String User_table_view
            {
                get
                {
                    return configUser_table_view;
                }
                set
                {
                    configUser_table_view = value;
                }
            }
            public String User_sAMAccount
            {
                get
                {
                    return configUser_sAMAccount;
                }
                set
                {
                    configUser_sAMAccount = value;
                }
            }
            public String User_dbTable
            {
                get
                {
                    return configUser_dbTable;
                }
                set
                {
                    configUser_dbTable = value;
                }
            }
            public String User_where
            {
                get
                {
                    return configUser_where;
                }
                set
                {
                    configUser_where = value;
                }
            }
            public String DataServer
            {
                get
                {
                    return configDataServer;
                }
                set
                {
                    configDataServer = value;
                }
            }
            public String DBCatalog
            {
                get
                {
                    return configDBCatalog;
                }
                set
                {
                    configDBCatalog = value;
                }
            }

            // creates a dictionay of values
            public Dictionary<string, string> ToDictionary()
            {
                Dictionary<string, string> returnvalue = new Dictionary<string, string>();
                returnvalue.Add("configBaseGroupOU", configBaseGroupOU);
                returnvalue.Add("configBaseUserOU", configBaseUserOU);
                returnvalue.Add("configNotes", configNotes);
                returnvalue.Add("configGroup_CN", configGroup_CN);
                returnvalue.Add("configGroup_table_view", configGroup_table_view);
                returnvalue.Add("configGroup_sAMAccount", configGroup_sAMAccount);
                returnvalue.Add("configGroup_dbTable", configGroup_dbTable);
                returnvalue.Add("configGroup_where", configGroup_where);
                returnvalue.Add("configGroup_Prepend", configGroup_Prepend);
                returnvalue.Add("configUser_Group_Reference", configUser_Group_Reference);
                returnvalue.Add("configUser_table_view", configUser_table_view);
                returnvalue.Add("configUser_sAMAccount", configUser_sAMAccount);
                returnvalue.Add("configUser_dbTable", configUser_dbTable);
                returnvalue.Add("configUser_where", configUser_where);
                returnvalue.Add("configDataServer", configDataServer);
                returnvalue.Add("configDBCatalog", configDBCatalog);
                return returnvalue;
            }
            public GroupSynch Clone()
            {
                GroupSynch retunvalue = new GroupSynch();
                retunvalue.Load(this.ToDictionary());
                return retunvalue;
            }

        }
        public class UserSynch
        {  
            private String configBaseUserOU;
            private String configUniversalGroup;
            private String configNotes;
            private String configUser_Lname;
            private String configUser_Fname;
            private String configUser_city;
            private String configUser_State;
            private String configUser_Zip;
            private String configUser_Address;
            private String configUser_Mobile;
            private String configUser_sAMAccount;
            private String configUser_table_view;
            private String configUser_dbTable;
            private String configUser_where;
            private String configDataServer;
            private String configDBCatalog;
            private String configUserHoldingTank;
            private String configUser_password;
            private DataTable configCustoms = new DataTable();
            private string custom = "";
            int i = 0;
            int j = 0;
            private DataRow row;



            // constructor creates a blank instance
            public UserSynch()
            {
                configUniversalGroup = "";
                configBaseUserOU = "";
                configNotes = "";
                configUser_Lname = "";
                configUser_Fname = "";
                configUser_city = "";
                configUser_State = "";
                configUser_Zip = "";
                configUser_Address = "";
                configUser_Mobile = "";
                configUser_table_view = "";
                configUser_sAMAccount = "";
                configUser_dbTable = "";
                configUser_where = "";
                configDataServer = "";
                configDBCatalog = "";
                configUserHoldingTank = "";
                configUser_password = "";
                configCustoms.Columns.Add("ad");
                configCustoms.Columns.Add("sql");
                configCustoms.Columns.Add("static");
            }
            public void load(Dictionary<string, string> dictionary)
            {
                dictionary.TryGetValue("configBaseUserOU", out configBaseUserOU);
                dictionary.TryGetValue("configUniversalGroup", out configUniversalGroup);
                dictionary.TryGetValue("configNotes", out configNotes);
                dictionary.TryGetValue("configUser_Lname", out configUser_Lname);
                dictionary.TryGetValue("configUser_Fname", out configUser_Fname);
                dictionary.TryGetValue("configUser_city", out configUser_city);
                dictionary.TryGetValue("configUser_Zip", out configUser_Zip);
                dictionary.TryGetValue("configUser_State", out configUser_State);
                dictionary.TryGetValue("configUser_Address", out configUser_Address);
                dictionary.TryGetValue("configUser_Mobile", out configUser_Mobile);
                dictionary.TryGetValue("configUser_table_view", out configUser_table_view);
                dictionary.TryGetValue("configUser_sAMAccount", out configUser_sAMAccount);
                dictionary.TryGetValue("configUser_dbTable", out configUser_dbTable);
                dictionary.TryGetValue("configUser_where", out configUser_where);
                dictionary.TryGetValue("configDataServer", out configDataServer);
                dictionary.TryGetValue("configDBCatalog", out configDBCatalog);
                dictionary.TryGetValue("configUserHoldingTank", out configUserHoldingTank);
                dictionary.TryGetValue("configUser_password", out configUser_password);
                dictionary.TryGetValue("configCustoms", out custom);

                configCustoms.Clear();
                row = configCustoms.NewRow();
                if (custom != "" && custom != null)
                {
                    string[] rows = custom.Split('&');
                    for (i = 0; i < rows.Length; i++)
                    {
                        string[] parts = rows[i].Split('^');
                        row[0] = parts[0].Trim();
                        row[1] = parts[1].Trim();
                        row[2] = parts[2].Trim();
                        configCustoms.Rows.Add(row);
                        row = configCustoms.NewRow(); 
                    }
                }
            }

            // accessors for properties
            public String User_Fname
            {
                get
                {
                    return configUser_Fname;
                }
                set
                {
                    configUser_Fname = value;
                }
            }
            public String User_city
            {
                get
                {
                    return configUser_city;
                }
                set
                {
                    configUser_city = value;
                }
            }
            public String User_State
            {
                get
                {
                    return configUser_State;
                }
                set
                {
                    configUser_State = value;
                }
            }
            public String User_Zip
            {
                get
                {
                    return configUser_Zip;
                }
                set
                {
                    configUser_Zip = value;
                }
            }
            public String User_Address
            {
                get
                {
                    return configUser_Address;
                }
                set
                {
                    configUser_Address = value;
                }
            }
            public String User_Mobile
            {
                get
                {
                    return configUser_Mobile;
                }
                set
                {
                    configUser_Mobile = value;
                }
            }
            public String BaseUserOU
            {
                get
                {
                    return configBaseUserOU;
                }
                set
                {
                    configBaseUserOU = value;
                }
            }
            public String Notes
            {
                get
                {
                    return configNotes;
                }
                set
                {
                    configNotes = value;
                }
            }
            public String User_Lname
            {
                get
                {
                    return configUser_Lname;
                }
                set
                {
                    configUser_Lname = value;
                }
            }
            public String User_table_view
            {
                get
                {
                    return configUser_table_view;
                }
                set
                {
                    configUser_table_view = value;
                }
            }
            public String User_sAMAccount
            {
                get
                {
                    return configUser_sAMAccount;
                }
                set
                {
                    configUser_sAMAccount = value;
                }
            }
            public String User_dbTable
            {
                get
                {
                    return configUser_dbTable;
                }
                set
                {
                    configUser_dbTable = value;
                }
            }
            public String User_where
            {
                get
                {
                    return configUser_where;
                }
                set
                {
                    configUser_where = value;
                }
            }
            public String DataServer
            {
                get
                {
                    return configDataServer;
                }
                set
                {
                    configDataServer = value;
                }
            }
            public String DBCatalog
            {
                get
                {
                    return configDBCatalog;
                }
                set
                {
                    configDBCatalog = value;
                }
            }
            public String UserHoldingTank
            {
                get
                {
                    return configUserHoldingTank;
                }
                set
                {
                    configUserHoldingTank = value;
                }
            }
            public String UniversalGroup
            {
                get
                {
                    return configUniversalGroup;
                }
                set
                {
                    configUniversalGroup = value;
                }
            }
            public String User_password
            {
                get
                {
                    return configUser_password;
                }
                set
                {
                    configUser_password = value;
                }
            }
            public DataTable UserCustoms
            {
                get
                {
                    return configCustoms;
                }
                set
                {
                    configCustoms = value;
                }
            }
            public String CustomsString
            {
                get
                {
                    return custom;
                }
                set
                {
                    custom = value;
                }
            }

            // output to a dictionary list
            public Dictionary<string, string> ToDictionary()
            {
                Dictionary<string, string> returnvalue = new Dictionary<string, string>();
                returnvalue.Add("configBaseUserOU", configBaseUserOU);
                returnvalue.Add("configNotes", configNotes);
                returnvalue.Add("configUser_Lname", configUser_Lname);
                returnvalue.Add("configUser_Fname", configUser_Fname);
                returnvalue.Add("configUser_city", configUser_city);
                returnvalue.Add("configUser_State", configUser_State);
                returnvalue.Add("configUser_Zip", configUser_Zip);
                returnvalue.Add("configUser_Address", configUser_Address);
                returnvalue.Add("configUser_Mobile", configUser_Mobile);
                returnvalue.Add("configUser_table_view", configUser_table_view);
                returnvalue.Add("configUser_sAMAccount", configUser_sAMAccount);
                returnvalue.Add("configUser_dbTable", configUser_dbTable);
                returnvalue.Add("configUser_where", configUser_where);
                returnvalue.Add("configDataServer", configDataServer);
                returnvalue.Add("configDBCatalog", configDBCatalog);
                returnvalue.Add("configUniversalGroup", configUniversalGroup);
                returnvalue.Add("configUserHoldingTank", configUserHoldingTank);
                returnvalue.Add("configUser_password", configUser_password);
                //for (i = 0; i < configCustoms.Rows.Count; i++)
                //{
                //    for (j = 0; j < configCustoms.Columns.Count; j++)
                //    {
                //        custom += configCustoms.Rows[i][j].ToString();
                //        custom += "^";
                //    }
                //    custom += "&";
                //}

                returnvalue.Add("configCustoms", custom);
                return returnvalue;
            }
        }
        public class UserStateChange
        {
            private String configDBorAD;
            // DB values
            private String configUser_CN;
            private String configUser_table_view;
            private String configUser_sAMAccount;
            private String configUser_dbTable;
            private String configUser_where;
            // AD selection values            
            private String configADOU;
            private String configADGroups;

            private String configActiveDisabled;
            private String configMoveOrDelete;
            private String configDisableTime;
            private String configFromOU;
            private String configToOU;

            // base ou to check for users in
            private String configBaseOU;
            private String configDataServer;
            private String configDBCatalog;

            // constructor creates a blank instance
            public UserStateChange()
            {
                configUser_CN = "";
                configUser_table_view = "";
                configUser_sAMAccount = "";
                configUser_dbTable = "";
                configUser_where = "";
                configDataServer = "";
                configDBCatalog = "";
                configDBorAD = "";
                configADOU = "";
                configADGroups = "";
                configMoveOrDelete = "";
                configBaseOU = "";
                configDisableTime = "";
                configFromOU = "";
                configToOU = "";
                configActiveDisabled = "";
            }
            public void load(Dictionary<string, string> dictionary)
            {
                dictionary.TryGetValue("configUser_CN", out configUser_CN);
                dictionary.TryGetValue("configUser_table_view", out configUser_table_view);
                dictionary.TryGetValue("configUser_sAMAccount", out configUser_sAMAccount);
                dictionary.TryGetValue("configUser_dbTable", out configUser_dbTable);
                dictionary.TryGetValue("configUser_where", out configUser_where);
                dictionary.TryGetValue("configDataServer", out configDataServer);
                dictionary.TryGetValue("configDBCatalog", out configDBCatalog);
                dictionary.TryGetValue("configDBorAD", out configDBorAD);
                dictionary.TryGetValue("configADOU", out configADOU);
                dictionary.TryGetValue("configADGroups", out configADGroups);
                dictionary.TryGetValue("configMoveOrDelete", out configMoveOrDelete);
                dictionary.TryGetValue("configBaseOU", out configBaseOU);
                dictionary.TryGetValue("configDisableTime", out configDisableTime);
                dictionary.TryGetValue("configFromOU", out configFromOU);
                dictionary.TryGetValue("configToOU", out configToOU);
                dictionary.TryGetValue("configActiveDisabled", out configActiveDisabled);
            }

            public String User_CN
            {
                get
                {
                    return configUser_CN;
                }
                set
                {
                    configUser_CN = value;
                }
            }
            public String User_table_view
            {
                get
                {
                    return configUser_table_view;
                }
                set
                {
                    configUser_table_view = value;
                }
            }
            public String User_sAMAccount
            {
                get
                {
                    return configUser_sAMAccount;
                }
                set
                {
                    configUser_sAMAccount = value;
                }
            }
            public String User_dbTable
            {
                get
                {
                    return configUser_dbTable;
                }
                set
                {
                    configUser_dbTable = value;
                }
            }
            public String User_where
            {
                get
                {
                    return configUser_where;
                }
                set
                {
                    configUser_where = value;
                }
            }
            public String DataServer
            {
                get
                {
                    return configDataServer;
                }
                set
                {
                    configDataServer = value;
                }
            }
            public String DBCatalog
            {
                get
                {
                    return configDBCatalog;
                }
                set
                {
                    configDBCatalog = value;
                }
            }
            public String ADOU
            {
                get
                {
                    return configADOU;
                }
                set
                {
                    configADOU = value;
                }
            }
            public String ADGroups
            {
                get
                {
                    return configADGroups;
                }
                set
                {
                    configADGroups = value;
                }
            }
            public String MoveOrDelete
            {
                get
                {
                    return configMoveOrDelete;
                }
                set
                {
                    configMoveOrDelete = value;
                }
            }
            public String BaseOU
            {
                get
                {
                    return configBaseOU;
                }
                set
                {
                    configBaseOU = value;
                }
            }
            public String DisableTime
            {
                get
                {
                    return configDisableTime;
                }
                set
                {
                    configDisableTime = value;
                }
            }
            public String FromOU
            {
                get
                {
                    return configFromOU;
                }
                set
                {
                    configFromOU = value;
                }
            }
            public String ToOU
            {
                get
                {
                    return configToOU;
                }
                set
                {
                    configToOU = value;
                }
            }
            public String ActiveDisabled
            {
                get
                {
                    return configActiveDisabled;
                }
                set
                {
                    configActiveDisabled = value;
                }
            }

            // output to array list
            public Dictionary<string, string> ToDictionary()
            {
                Dictionary<string, string> returnvalue = new Dictionary<string, string>();
                returnvalue.Add("configUser_CN", configUser_CN);
                returnvalue.Add("configUser_table_view", configUser_table_view);
                returnvalue.Add("configUser_sAMAccount", configUser_sAMAccount);
                returnvalue.Add("configUser_dbTable", configUser_dbTable);
                returnvalue.Add("configUser_where", configUser_where);
                returnvalue.Add("configDataServer", configDataServer);
                returnvalue.Add("configDBCatalog", configDBCatalog);
                returnvalue.Add("configDBorAD", configDBorAD);
                returnvalue.Add("configADGroups", configADGroups);
                returnvalue.Add("configMoveOrDelete", configMoveOrDelete);
                returnvalue.Add("configBaseOU", configBaseOU);
                returnvalue.Add("configDisableTime", configDisableTime);
                returnvalue.Add("configFromOU", configFromOU);
                returnvalue.Add("configToOU", configToOU);
                returnvalue.Add("configActiveDisabled", configActiveDisabled);
                return returnvalue;
            }

        }
        public class executionOrder
        {
            private ArrayList order;
            public ArrayList execution_order
            {
                get
                {
                    return order;
                }
                set
                {
                    order = value;
                }

            }
        }
        public class StopWatch
        {

            private DateTime startTime;
            private DateTime stopTime;
            private bool running = false;


            public void Start()
            {
                this.startTime = DateTime.Now;
                this.running = true;
            }


            public void Stop()
            {
                this.stopTime = DateTime.Now;
                this.running = false;
            }


            // elaspsed time in milliseconds
            public double GetElapsedTime()
            {
                TimeSpan interval;

                if (running)
                    interval = DateTime.Now - startTime;
                else
                    interval = stopTime - startTime;

                return interval.TotalMilliseconds;
            }


            // elaspsed time in seconds
            public double GetElapsedTimeSecs()
            {
                TimeSpan interval;

                if (running)
                    interval = DateTime.Now - startTime;
                else
                    interval = stopTime - startTime;

                return interval.TotalSeconds;
            }
        }
        public class ToolSet
        {

            //Functions
            public string GetDomain()
            {
                using (Domain d = Domain.GetCurrentDomain())
                using (DirectoryEntry entry = d.GetDirectoryEntry())
                {
                    return entry.Path;
                }
            }
            public DataTable EnumerateUsersInOUDataTable(string ouDN, ArrayList returnProperties, string table)
            {
                // note does not handle special/illegal characters for AD
                // RETURNS ALL USERS IN AN OU NO MATTER HOW DEEP
                int count = returnProperties.Count;
                int i;
                DataTable returnvalue = new DataTable();
                DataRow row;
                

                // bind to the OU you want to enumerate
                DirectoryEntry deOU = new DirectoryEntry("LDAP://" + ouDN);

                // create a directory searcher for that OU
                DirectorySearcher dsUsers = new DirectorySearcher(deOU);

                // set depth to recursive
                dsUsers.SearchScope = SearchScope.Subtree;

                // set the filter to get just the users
                dsUsers.Filter = "(&(objectClass=user)(objectCategory=Person))";

                // add the attributes you want to grab from the search
                // add the attributes you want to grab from the search
                for (i = 0; i < count; i++)
                {
                    dsUsers.PropertiesToLoad.Add(returnProperties[i].ToString());
                    returnvalue.Columns.Add(returnProperties[i].ToString());
                } 
                //dsUsers.PropertiesToLoad.Add("sAMAccountName");

                // grab the users and do whatever you need to do with them 
                dsUsers.PageSize = 500;
                row = returnvalue.NewRow();
                foreach (SearchResult oResult in dsUsers.FindAll())
                {
                    //generate the array list with the user sam accounts
                    for (i = 0; i < count; i++)
                    {
                        try
                        {
                            row[i] = System.Web.HttpUtility.UrlDecode(oResult.Properties[returnProperties[i].ToString()][0].ToString());
                        }
                        catch
                        {
                            row[i] = string.Empty;
                        }
                    }
                    returnvalue.Rows.Add(row);
                    row = returnvalue.NewRow();
                }
                return returnvalue;
            }
            public DataTable EnumerateUsersInGroupDataTable(string groupDN, string groupou, string coulumnNameForFQDN, string coulumnNameForGroup, string table)
            {
                // note does not handle special/illegal characters for AD
                // it is optimal to have the two field names match the coumn names pulled from SQL do not use group fro group name it will kill the SQL
                // groupDN "CN=Sales,OU=test,DC=Fabrikam,DC=COM"
                // returns FQDN "CN=user,OU=test,DC=Fabrikam,DC=COM" & group "CN=Sales,OU=test,DC=Fabrikam,DC=COM" of users in group 
                DataTable returnvalue = new DataTable();
                DataRow row;

                string bladh = "LDAP://" + "CN=" + System.Web.HttpUtility.UrlEncode(groupDN).Replace("+", " ").Replace("*", "%2A") + groupou;
                DirectoryEntry group = new DirectoryEntry("LDAP://" + "CN=" + System.Web.HttpUtility.UrlEncode(groupDN).Replace("+", " ").Replace("*", "%2A") + groupou);
                DirectorySearcher groupUsers = new DirectorySearcher(group);
                returnvalue.TableName = table;
                returnvalue.Columns.Add(coulumnNameForFQDN);
                returnvalue.Columns.Add(coulumnNameForGroup);
                row = returnvalue.NewRow();
                foreach (object dn in group.Properties["member"])
                {
                    row[0] = dn.ToString();
                    row[1] = groupDN;                    
                    returnvalue.Rows.Add(row);
                    row = returnvalue.NewRow();
                }
                return returnvalue;
            }
            public DataTable EnumerateGroupsInOUDataTable(string ouDN, ArrayList returnProperties, string table)
            {
                // note does not handle special/illegal characters for AD
                int i;
                int count = returnProperties.Count;
                DataTable returnvalue = new DataTable();
                DataRow row;
                // bind to the OU you want to enumerate
                DirectoryEntry deOU = new DirectoryEntry("LDAP://" + ouDN);                

                // create a directory searcher for that OU
                DirectorySearcher dsUsers = new DirectorySearcher(deOU);

                // set the filter to get just the users
                dsUsers.Filter = "(&(objectClass=group))";
                // make it non recursive in depth
                dsUsers.SearchScope = SearchScope.OneLevel;

                returnvalue.TableName = table;
                // add the attributes you want to grab from the search
                for (i = 0; i < count; i++)
                {
                    dsUsers.PropertiesToLoad.Add(returnProperties[i].ToString());
                    returnvalue.Columns.Add(returnProperties[i].ToString());
                }                
                // grab the users and do whatever you need to do with them 
                dsUsers.PageSize = 500;
                row = returnvalue.NewRow();
                foreach (SearchResult oResult in dsUsers.FindAll())
                {
                    //generate the array list with the user sam accounts
                    for (i = 0; i < count; i++)
                    {
                        try
                        {
                            row[i] = System.Web.HttpUtility.UrlDecode(oResult.Properties[returnProperties[i].ToString()][0].ToString());
                        }
                        catch
                        {
                            row[i] = string.Empty;
                        }
                    }
                    returnvalue.Rows.Add(row);
                    row = returnvalue.NewRow();
                }
                return returnvalue;
            }
            public bool Exists(string objectPath)
            {
                // checks if there is a object for the distinguished name
                bool found = false;
                if ((objectPath != null) && (objectPath != string.Empty))
                {
                    try
                    {
                        if (DirectoryEntry.Exists("LDAP://" + objectPath))
                        {
                            found = true;
                        }
                    }
                    catch (Exception e)
                    {
                        // MessageBox.Show(e.Message.ToString() + "create group LDAP://CN=" + name + "," + ouPath);
                        return found;
                    }
                }
                return found;
            }
            public string AttributeValuesSingleString(string attributeName, string objectDn)
            {
                string strValue;
                DirectoryEntry ent = new DirectoryEntry(objectDn);
                strValue = ent.Properties[attributeName].Value.ToString();
                ent.Close();
                ent.Dispose();
                return strValue;
            }
            public string GetObjectDistinguishedName(objectClass objectCls, returnType returnValue, string objectName, string LdapDomain)
            {
                // LdapDomain = "DC=Fabrikam,DC=COM" 

                string distinguishedName = string.Empty;
                string connectionPrefix = "LDAP://" + LdapDomain;
                DirectoryEntry entry = new DirectoryEntry(connectionPrefix);
                DirectorySearcher mySearcher = new DirectorySearcher(entry);

                switch (objectCls)
                {
                    case objectClass.user:
                        mySearcher.Filter = "(&(objectClass=user)(|(CN=" + objectName + ")(sAMAccountName=" + objectName + ")))";
                        break;
                    case objectClass.group:
                        mySearcher.Filter = "(&(objectClass=group)(|(CN=" + objectName + ")(dn=" + objectName + ")))";
                        break;
                    case objectClass.computer:
                        mySearcher.Filter = "(&(objectClass=computer)(|(CN=" + objectName + ")(dn=" + objectName + ")))";
                        break;
                }
                SearchResult result = mySearcher.FindOne();

                if (result == null)
                {
                    //throw new NullReferenceException
                    //("unable to locate the distinguishedName for the object " +
                    //objectName + " in the " + LdapDomain + " domain");
                    return string.Empty;
                }
                DirectoryEntry directoryObject = result.GetDirectoryEntry();
                if (returnValue.Equals(returnType.distinguishedName))
                {
                    distinguishedName = "LDAP://" + directoryObject.Properties
                        ["distinguishedName"].Value;
                }
                if (returnValue.Equals(returnType.ObjectGUID))
                {
                    distinguishedName = directoryObject.Guid.ToString();
                }
                entry.Close();
                entry.Dispose();
                mySearcher.Dispose();
                return distinguishedName;
            }
            public bool CreateOURecursive(string ou, LogFile log)
            {
                try
                {
                    if (Exists(ou) == true)
                    {
                        return true;
                    }
                    else
                    {
                        CreateOURecursive(ou.Substring(ou.IndexOf(",") + 1), log);
                        CreateOU(ou.Substring(ou.IndexOf(",") + 1), ou.Remove(ou.IndexOf(",")).Substring(ou.IndexOf("=") + 1), log);
                        return true;
                    }
                }
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    log.errors.Add(E.Message.ToString() + " error creating OU" + ou);
                    return false;
                }
            }
            public void CreateGroup(string ouPath, Dictionary<string, string> properties, LogFile log)
            {
                // otherProperties is a mapping  <the key is the active driectory field, and the value is the the value>
                // the keys must contain valid AD fields
                // the value will relate to the specific key
                //needs parent OU present to work
                try
                {
                    if (!DirectoryEntry.Exists("LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath))
                    {

                        DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                        DirectoryEntry group = entry.Children.Add("CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A"), "group");
                        foreach (KeyValuePair<string, string> kvp in properties)
                        {
                             group.Properties[kvp.Key.ToString()].Value = System.Web.HttpUtility.UrlEncode(kvp.Value.ToString()).Replace("+", " ").Replace("*", "%2A");                            
                        }
                        group.CommitChanges();
                        group.Close();
                        group.Dispose();
                        entry.Close();
                        entry.Dispose();
                        log.transactions.Add("group added | LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath);
                    }
                    else
                    { 
                        log.warnings.Add("CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + " group already exists from adding");
                    }
                }
                catch (Exception e)
                {
                    log.errors.Add(e.Message.ToString() + "issue create group LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath);
                }
            }
            public void UpdateGroup(string ouPath, Dictionary<string, string> properties, LogFile log)
            {
                // otherProperties is a mapping  <the key is the active driectory field, and the value is the the value>
                // the keys must contain valid AD fields
                // the value will relate to the specific key
                // needs parent OU present to work
                if (DirectoryEntry.Exists("LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath))
                {
                    try
                    {
                        DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                        DirectoryEntry group = entry.Children.Find("CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A"));
                        foreach (KeyValuePair<string, string> kvp in properties)
                        {
                            if (kvp.Key.ToString() == "CN" || kvp.Key.ToString() == "sAMAccountName")
                            { }
                            else
                            {
                                group.Properties[kvp.Key.ToString()].Value = System.Web.HttpUtility.UrlEncode(kvp.Value.ToString()).Replace("+", " ").Replace("*", "%2A");
                            }
                        }
                        group.CommitChanges();
                        group.Close();
                        group.Dispose();
                        entry.Close();
                        entry.Dispose();
                        log.transactions.Add("updated group | LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath);
                    }
                    catch (Exception e)
                    {
                        log.errors.Add(e.Message.ToString() + "issue updating group LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath);
                    }
                }
                else
                { 
                    log.warnings.Add(ouPath + " group does not exist");
                }
            }
            public void DeleteGroup(string ouPath, string name, LogFile log)
            {
                if (DirectoryEntry.Exists("LDAP://CN=" + System.Web.HttpUtility.UrlEncode(name).Replace("+", " ").Replace("*", "%2A") + "," + ouPath))
                {
                    try
                    {
                        DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                        DirectoryEntry group = new DirectoryEntry("LDAP://CN=" + System.Web.HttpUtility.UrlEncode(name).Replace("+", " ").Replace("*", "%2A") + "," + ouPath);
                        entry.Children.Remove(group);
                        group.CommitChanges();
                        group.Close();
                        group.Dispose();
                        entry.Close();
                        entry.Dispose();
                        log.transactions.Add("deleted group | LDAP://CN=" + name + "," + ouPath);
                    }
                    catch (Exception e)
                    {
                        log.errors.Add(e.Message.ToString() + " error deleting LDAP://CN=" + name + "," + ouPath );
                    }
                }
                else
                {
                    log.warnings.Add("group LDAP://CN=" + name + "," + ouPath + " does not exists cannot delete" );
                }
            }
            public void CreateOU(string ouPath, string name, LogFile log)
            {
                //needs parent OU present to work
                if (!DirectoryEntry.Exists("LDAP://OU=" + name + "," + ouPath))
                {
                    try
                    {
                        DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                        DirectoryEntry OU = entry.Children.Add("OU=" + name, "organizationalUnit");
                        OU.CommitChanges();
                        OU.Close();
                        OU.Dispose();
                        entry.Close();
                        entry.Dispose();
                        log.transactions.Add("created ou | LDAP://OU=" + name + "," + ouPath);
                    }
                    catch (Exception e)
                    {
                        log.errors.Add(e.Message.ToString() + "error creating ou LDAP://OU=" + name + "," + ouPath);
                    }
                }
                else
                { 
                    log.warnings.Add("creating ou LDAP://OU=" + name + "," + ouPath + " already exists"); 
                }
            }
            public void DeleteOU(string ouPath, string name, LogFile log)
            {
                //needs parent OU present to work
                if (!DirectoryEntry.Exists("LDAP://OU=" + name + "," + ouPath))
                {
                    try
                    {
                        DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                        entry.DeleteTree();
                        entry.Close();
                        entry.Dispose();
                        log.transactions.Add("deleting ou | LDAP://OU=" + name + "," + ouPath + " does not exists");
                    }
                    catch (Exception e)
                    {
                        log.errors.Add(e.Message.ToString() + " error deleting ou LDAP://OU=" + name + "," + ouPath);
                    }
                }
                else
                { 
                    log.warnings.Add("error deleting ou LDAP://OU=" + name + "," + ouPath + " does not exists"); 
                }
            }
            public void AddUserToGroup(string userDn, string groupDn, LogFile log)
            {
                try
                {
                    if (Exists(userDn) && Exists(groupDn))
                    {
                        DirectoryEntry entry = new DirectoryEntry("LDAP://" + groupDn);
                        entry.Properties["member"].Add(userDn);
                        // dirEntry.Invoke("Add", new object[] { "LDAP://" + userDn });
                        entry.CommitChanges();
                        entry.Close();
                        entry.Dispose();
                        log.transactions.Add("added user to group | " + userDn + " | LDAP://" + groupDn);
                    }
                    else
                    {                
                    log.warnings.Add(" Warning could not add user " + userDn + " to group LDAP://" + groupDn + " group did not exist");

                    }

                }
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    log.errors.Add(E.Message.ToString() + " error adding user to group" + userDn + " to LDAP://" + groupDn);

                }
            }
            public void RemoveUserFromGroup(string userDn, string groupDn, LogFile log)
            {
                try
                {
                    DirectoryEntry entry = new DirectoryEntry("LDAP://" + groupDn);
                    try
                    {
                        entry.Properties["member"].Remove(userDn);
                        log.transactions.Add("removed user from group | " + userDn + " | LDAP://" + groupDn);
                    }
                    catch (System.DirectoryServices.DirectoryServicesCOMException E)
                    {
                        log.errors.Add(E.Message.ToString() + " error removing user from group " + userDn + " from LDAP://" + groupDn + " user may not be in group");
                    }
                    entry.CommitChanges();
                    entry.Close();
                    entry.Dispose();
                }
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    log.warnings.Add(E.Message.ToString() + " error removing user from group " + userDn + " from LDAP://" + groupDn + " group object does not exist");
                }
            }


            public void CreateUserAccount(string ouPath, SqlDataReader users, string groupDn, UserSynch usersyn, LogFile log)
            {
                // properties contians key pairs with the key matching a field within Active driectory and the value, the value to be inserted
                int i;                
                int fieldcount;
                int val; 
                string name = "";
                fieldcount = users.FieldCount;
                while (users.Read())
                {
                    try
                    {
                        if (users[usersyn.User_password].ToString() != "")
                        {
                            if (!DirectoryEntry.Exists("LDAP://CN=" + System.Web.HttpUtility.UrlEncode(users[usersyn.User_sAMAccount].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath))
                            {

                                DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                                DirectoryEntry newUser = entry.Children.Add("CN=" + System.Web.HttpUtility.UrlEncode(users[usersyn.User_sAMAccount].ToString()).Replace("+", " ").Replace("*", "%2A"), "user");
                                 // generated
                                newUser.Properties["samAccountName"].Value = System.Web.HttpUtility.UrlEncode(users[usersyn.User_sAMAccount].ToString()).Replace("+", " ").Replace("*", "%2A");
                                newUser.Properties["UserPrincipalName"].Value = System.Web.HttpUtility.UrlEncode(users[usersyn.User_sAMAccount].ToString()).Replace("+", " ").Replace("*", "%2A");
                                newUser.Properties["displayName"].Value = System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Lname]).Replace("+", " ").Replace("*", "%2A") + ", " + System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Fname]).Replace("+", " ").Replace("*", "%2A");
                                newUser.Properties["description"].Value = System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Lname]).Replace("+", " ").Replace("*", "%2A") + ", " + System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Fname]).Replace("+", " ").Replace("*", "%2A");
                                newUser.CommitChanges();

                                // SQL query generated ensures matching field names between the SQL form fields and AD
                                for (i = 0; i < fieldcount; i++)
                                {
                                    name = users.GetName(i);
                                    if (name != "password" && name != "CN")
                                    {
                                        if ((string)users[i] != "")
                                        {
                                            newUser.Properties[users.GetName(i)].Value = System.Web.HttpUtility.UrlEncode((string)users[i]).Replace("+", " ").Replace("*", "%2A");
                                        }
                                    }
                                }   

                                
                                AddUserToGroup("CN=" + System.Web.HttpUtility.UrlEncode(users[usersyn.User_sAMAccount].ToString()).Replace("+", " ").Replace("*", "%2A") +  "," + usersyn.UserHoldingTank, groupDn, log);
                                newUser.Invoke("SetPassword", new object[] { System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_password]).Replace("+", " ").Replace("*", "%2A") });
                                newUser.CommitChanges();

                                val = (int)newUser.Properties["userAccountControl"].Value;
                                // set to normal user
                                newUser.Properties["userAccountControl"].Value = val | (int)accountFlags.ADS_UF_NORMAL_ACCOUNT;
                                // set to enabled account val & ~0c0002 creates a bitmask which reverses the disabled bit
                                newUser.Properties["userAccountControl"].Value = val & ~(int)accountFlags.ADS_UF_ACCOUNTDISABLE;
                                newUser.CommitChanges();
                                newUser.Close();
                                newUser.Dispose();
                                entry.Close();
                                entry.Dispose();
                                log.transactions.Add("User added |" + (string)users[usersyn.User_sAMAccount] + " " + usersyn.UserHoldingTank);
                            }
                            else
                            {
                                log.errors.Add("CN=" + System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_sAMAccount]).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + " user already exists from adding");
                                //MessageBox.Show("CN=" + System.Web.HttpUtility.UrlEncode((string)users["CN"]).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + " user already exists from adding");
                            }
                        }

                    }
                    catch (Exception e)
                    {
                        string debugdata = "";
                        for (i = 0; i < fieldcount; i++)
                        {

                          debugdata += users.GetName(i) + "=" + System.Web.HttpUtility.UrlEncode((string)users[i]).Replace("+", " ").Replace("*", "%2A") + ", ";

                        }
                        log.errors.Add("issue create user LDAP://CN=" + System.Web.HttpUtility.UrlEncode((string)users["CN"]).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + "\n" + debugdata +  " failed field maybe " + name + " | " + e.Message.ToString());
                        // MessageBox.Show(e.Message.ToString() + "issue create user LDAP://CN=" + System.Web.HttpUtility.UrlEncode((string)users["CN"]).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + "\n" + debugdata);
                    }
                }

                /*
                    //Add this to the create account method
                    int val = (int)newUser.Properties["userAccountControl"].Value; 
                         //newUser is DirectoryEntry object
                    newUser.Properties["userAccountControl"].Value = val | 0x80000; 
                        //ADS_UF_TRUSTED_FOR_DELEGATION
                 
                 * 
                 * UserAccountControlFlags
                 * CONST   HEX
                    -------------------------------
                    SCRIPT 0x0001
                    ACCOUNTDISABLE 0x0002
                    HOMEDIR_REQUIRED 0x0008
                    LOCKOUT 0x0010
                    PASSWD_NOTREQD 0x0020
                    PASSWD_CANT_CHANGE 0x0040
                    ENCRYPTED_TEXT_PWD_ALLOWED 0x0080
                    TEMP_DUPLICATE_ACCOUNT 0x0100
                    NORMAL_ACCOUNT 0x0200
                    INTERDOMAIN_TRUST_ACCOUNT 0x0800
                    WORKSTATION_TRUST_ACCOUNT 0x1000
                    SERVER_TRUST_ACCOUNT 0x2000
                    DONT_EXPIRE_PASSWORD 0x10000
                    MNS_LOGON_ACCOUNT 0x20000
                    SMARTCARD_REQUIRED 0x40000
                    TRUSTED_FOR_DELEGATION 0x80000
                    NOT_DELEGATED 0x100000
                    USE_DES_KEY_ONLY 0x200000
                    DONT_REQ_PREAUTH 0x400000
                    PASSWORD_EXPIRED 0x800000
                    TRUSTED_TO_AUTH_FOR_DELEGATION 0x1000000
                 * */
            }
            public bool DisableUser(string sAMAccountName, string ldapDomain, LogFile log)
            {
                string userDN;
                userDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, ldapDomain);
                try
                {  
                    DirectoryEntry usr = new DirectoryEntry(userDN);
                    int val = (int)usr.Properties["userAccountControl"].Value;
                    usr.Properties["userAccountControl"].Value = val | (int)accountFlags.ADS_UF_ACCOUNTDISABLE;
                    usr.CommitChanges();
                    usr.Close();
                    usr.Dispose();
                    log.transactions.Add("diabled user account |" + userDN);
                    return true;
                }
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    log.errors.Add(E.Message.ToString() + " error disabling user " + userDN);
                    return false;
                }
                
            }
            public bool EnableUser(string sAMAccountName, string ldapDomain, LogFile log)
            {
                string userDN;
                userDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, ldapDomain);
                try
                { 
                    DirectoryEntry usr = new DirectoryEntry(userDN);
                    int val = (int)usr.Properties["userAccountControl"].Value;
                    usr.Properties["userAccountControl"].Value = val | ~(int)accountFlags.ADS_UF_ACCOUNTDISABLE;
                    usr.CommitChanges();
                    usr.Close();
                    usr.Dispose();
                    log.transactions.Add("enabled user account |" + userDN);
                    return true;
                }
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    log.errors.Add(E.Message.ToString() + " error enabling user " + userDN);
                    return false;
                }
            }
            // additional stuff
            public bool SetUserExpiration(int days, string ldapDomain, string sAMAccountName, LogFile log)
            {
                string userDN;
                userDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, ldapDomain);
                try
                {
                    DirectoryEntry usr = new DirectoryEntry(userDN);
                    Type type = usr.NativeObject.GetType();
                    Object adsNative = usr.NativeObject;
                    string formattedDate;

                    

                    // Calculating the new date
                    DateTime yesterday = DateTime.Today.AddDays(days);
                    formattedDate = yesterday.ToString("dd/MM/yyyy");

                    type.InvokeMember("AccountExpirationDate", BindingFlags.SetProperty, null, adsNative, new object[] { formattedDate });
                    usr.CommitChanges();
                    usr.Close();
                    usr.Dispose();
                    log.transactions.Add("User expiration set |" + userDN + "|" + days);

                    return true;
                }
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    log.errors.Add(E.Message.ToString() + " error setting user expiration " + userDN + " days " + days);
                    return false;
                }
            }
            public bool SetUserExpiration(int days, string userDN, LogFile log)
            {
                try
                {
                    DirectoryEntry usr = new DirectoryEntry(userDN);
                    Type type = usr.NativeObject.GetType();
                    Object adsNative = usr.NativeObject;
                    string formattedDate;

                    // Calculating the new date
                    DateTime yesterday = DateTime.Today.AddDays(days);
                    formattedDate = yesterday.ToString("dd/MM/yyyy");

                    type.InvokeMember("AccountExpirationDate", BindingFlags.SetProperty, null, adsNative, new object[] { formattedDate });
                    usr.CommitChanges();
                    usr.Close();
                    usr.Dispose();
                    log.transactions.Add("User expiration set |" + userDN + "|" + days);
                    return true;
                }
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    log.errors.Add(E.Message.ToString() + " error setting user expiration " + userDN + " days " + days);
                    return false;
                }
            }
            public bool DeleteUserAccount(string sAMAccountName, string ldapDomain, LogFile log)
            {
                string userDN;
                userDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, ldapDomain);
                try
                { 
                    DirectoryEntry ent = new DirectoryEntry(userDN);
                    ent.DeleteTree();
                    ent.Close();
                    ent.Dispose();
                    log.transactions.Add("deleted user account |" + userDN);
                    return true;
                }
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    log.errors.Add(E.Message.ToString() + " error deleting user " + userDN);
                    return false;
                }
            }


            public string Create_Table(DataTable data, string table, SqlConnection sqlConn)
            {
                int i;
                int Count;
                StringBuilder sqlstring = new StringBuilder();
                SqlCommand sqlComm;
                Count = data.Columns.Count;

                // make the temp table
                sqlstring.Append("Create table " + table + "(");
                for (i = 0; i < Count; i++)
                {
                    sqlstring.Append(data.Columns[i] + " VarChar(350), ");
                }
                sqlstring.Remove((sqlstring.Length - 2), 2);
                sqlstring.Append(")");
                sqlComm = new SqlCommand(sqlstring.ToString(), sqlConn);
                try
                {
                    sqlComm.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "An Big poblem arose with the table create", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                }

                // copy data into table
                SqlBulkCopy sbc = new SqlBulkCopy(sqlConn);
                sbc.DestinationTableName = table;
                sbc.WriteToServer(data);
                sbc.Close();
                return table;
            }
            public string Append_to_Table(DataTable data, string table, SqlConnection sqlConn)
            {
                // copy data into table
                SqlBulkCopy sbc = new SqlBulkCopy(sqlConn);
                sbc.DestinationTableName = table;
                sbc.WriteToServer(data);
                sbc.Close();
                return table;
            }
            public SqlDataReader QueryNotExists(string table1, string table2, SqlConnection sqlConn, string pkey1, string pkey2)
            {
                // finds items in table1 who do not exist in table2 and returns them
                // SqlCommand sqlComm = new SqlCommand("Select Table1.* Into #Table3ADTransfer From " + Table1 + " AS Table1, " + Table2 + " AS Table2 Where Table1." + pkey1 + " = Table2." + pkey2 + " And Table2." + pkey2 + " is null", sqlConn);
                SqlCommand sqlComm = new SqlCommand("SELECT uptoDate.* FROM " + table1 + " uptoDate LEFT OUTER JOIN " + table2 + " outofDate ON outofDate." + pkey2 + " = uptoDate." + pkey1 + " WHERE outofDate." + pkey2 + " IS NULL;", sqlConn);
                // create the command object
                SqlDataReader r = sqlComm.ExecuteReader();
                return r;
            }
            public SqlDataReader CheckUpdate(string table1, string table2, string pkey1, string pkey2, ArrayList compareFields1, ArrayList compareFields2, SqlConnection sqlConn)
            {
                // Assumes table1 holds the correct data and returns a data reader with the update fields columns from table1
                // returns the rows which table2 differs from table1
                string compare1 = "";
                string compare2 = "";
                string fields = "";
                // need a comand builder and research on the best way to compare all fields in a row
                // this basically will just issue a concatenation sql query to the DB for each field to compare
                foreach (string key in compareFields1)
                {
                    compare1 = compare1 + table1 + "." + key + " + ";
                    fields += table1 + "." + key + ", ";
                }
                foreach (string key in compareFields2)
                {
                    compare2 = compare2 + table2 + "." + key + " + ";
                }
                // remove trailing comma and + 
                compare2 = compare2.Remove(compare2.Length - 2);
                compare1 = compare1.Remove(compare1.Length - 2);
                fields = fields.Remove(fields.Length - 2);
                SqlCommand sqlComm = new SqlCommand("SELECT " + fields + " FROM " + table1 + " INNER JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2 + " WHERE (" + compare2 + ") <> (" + compare1 + ")", sqlConn);
                //AND " + table2 + "." + pkey2 + " != NULL
                SqlDataReader r = sqlComm.ExecuteReader();
                return r;
            }
  
            public string SetAttributesForUser()
            {
                string returnvalue = "";
                //DirectoryEntry entry = null;

                //entry = new DirectoryEntry("LDAP://" + objectClass);

                //foreach (string propertyName in entry.Properties.PropertyNames)
                //{
                //    returnvalue += propertyName + "   :   " +
                //       entry.Properties[propertyName][0].ToString() 
                //       + " \n";
                    
                //}
                //return returnvalue;



                // bind to the OU you want to enumerate

                DirectoryEntry deOU = new DirectoryEntry(GetDomain());
                // create a directory searcher for that OU
                DirectorySearcher dsUsers = new DirectorySearcher(deOU);

                // set depth to recursive
                dsUsers.SearchScope = SearchScope.Subtree;

                // set the filter to get just the users
                dsUsers.Filter = "(&(objectClass=user)(objectCategory=Person))";

                SearchResult result = dsUsers.FindOne();                  

                if (result == null)
                {
                    //throw new NullReferenceException
                    //("unable to locate the distinguishedName for the object " +
                    //objectName + " in the " + LdapDomain + " domain");
                    return string.Empty;
                }
                DirectoryEntry user = result.GetDirectoryEntry();
                foreach (string propertyName in user.Properties.PropertyNames)
                {
                    returnvalue += propertyName + "   :   " +
                       user.Properties[propertyName][0].ToString()
                       + " \n";

                }
                return returnvalue;


                //DirectoryEntry mike = new DirectoryEntry("LDAP://schema/" + strGivenClass);
               // mike.Properties.PropertyNames
                             //Set objClass = GetObject("LDAP://schema/" & strGivenClass)

            //i = 0
            //For Each strPropName In objClass.MandatoryProperties
            //    i = i + 1
            //    WScript.Echo "Man " & i & ": " & strPropName
            //Next

            //i = 0
            //For Each strPropName In objClass.OptionalProperties
            //    i = i + 1
            //    WScript.Echo "Opt " & i & ": " & strPropName
            //Next
            }
            public ArrayList ADobjectAttribute()
            {
                // NOTE: One place where managed ADSI (System.DirectoryServices) falls short is finding schema 
                //information from LDAP/AD objects. Finding information like mandatory and optional
                //properties simply cannot be done with any managed classes

                DirectoryEntry schemaEntry = null;
                ArrayList returnvalue = new ArrayList();

                schemaEntry = new DirectoryEntry("LDAP://schema/user");
                ActiveDs.IADsClass iadsClass = (ActiveDs.IADsClass)schemaEntry.NativeObject;
                if (iadsClass == null)
                    return new ArrayList();


                ArrayList list = new ArrayList();
                foreach (string s in (Array)iadsClass.OptionalProperties)
                {
                    returnvalue.Add(s);
                }
                foreach (string s in (Array)iadsClass.MandatoryProperties)
                {
                    returnvalue.Add(s);
                }
                return returnvalue;

            }

            public ArrayList SqlColumns(UserSynch userconfig)
            {
                ArrayList columnList = new ArrayList();
                if (userconfig.DBCatalog != "" && userconfig.DataServer != "")
                {
                    //populates columns dialog with columns depending on the results of a query
                    
                    SqlConnection sqlConn = new SqlConnection("Data Source=" + userconfig.DataServer.ToString() + ";Initial Catalog=" + userconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");

                    sqlConn.Open();
                    // create the command object
                    SqlCommand sqlComm = new SqlCommand("SELECT column_name FROM information_schema.columns WHERE table_name = '" + userconfig.User_dbTable + "'", sqlConn);
                    SqlDataReader r = sqlComm.ExecuteReader();
                    while (r.Read())
                    {
                        columnList.Add((string)r[0].ToString().Trim());
                    }
                    r.Close();
                    sqlConn.Close();
                }
                return columnList;
            }
            // possible not in use

            public void MoveADObject(string objectLocation, string newLocation)
            {
                //For brevity, removed existence checks
                // EXPECTS FULL Distinguished Name for both variables "LDAP://CN=xxx,DC=xxx,DC=xxx"

                DirectoryEntry eLocation = new DirectoryEntry(objectLocation);
                DirectoryEntry nLocation = new DirectoryEntry(newLocation);
                string newName = eLocation.Name;
                eLocation.MoveTo(nLocation, newName);
                nLocation.Close();
                eLocation.Close();
            }
            public ArrayList SetMultiPropertyUser(LinkedList<Dictionary<string, string>> userList, ArrayList propertyArray, string ldapDomain, LogFile log)
            {
                /*
                 * takes a dictionary like such
                 * sAMAccountName, userName
                 * property1, value
                 * property2, value
                 * ....
                 * 
                 * and an arraylist of strings with the names of the keys for the properties
                 * ["property1", "property2", ...etc]
                 * 
                 * RETURNS
                 * users not found
                 */

                LinkedListNode<Dictionary<string, string>> userListNode = userList.First;
                string sAMAccountName;
                string userProperty;
                string usrDN;
                ArrayList returnvalue = new ArrayList();

                while (userListNode != null)
                {

                    userListNode.Value.TryGetValue("sAMAccountName", out sAMAccountName);
                    usrDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, ldapDomain);
                    if (usrDN == "")
                    {
                        returnvalue.Add(sAMAccountName);
                    }

                    //// get usr object for manipulation
                    DirectoryEntry user = new DirectoryEntry(usrDN);


                    foreach (string props in propertyArray)
                    {
                        userListNode.Value.TryGetValue(props, out userProperty);
                        user.Properties[props].Value = userProperty;
                    }
                    user.CommitChanges();
                    userListNode = userListNode.Next;
                }
                return returnvalue;
            }           
            public bool CreateUserAccount(string parentOUDN, string samName, string userPassword, string firstName, string lastName)
            {
                try
                {
                    string connectionPrefix = "LDAP://" + parentOUDN;
                    DirectoryEntry de = new DirectoryEntry(connectionPrefix);
                    DirectoryEntry newUser = de.Children.Add("CN=" + firstName + " " + lastName, "user");
                    newUser.Properties["samAccountName"].Value = samName;
                    newUser.Properties["userPrincipalName"].Value = samName;
                    newUser.Properties["sn"].Add(lastName);
                    newUser.Properties["name"].Value = firstName + " " + lastName;
                    newUser.Properties["givenName"].Add(firstName);

                    //foreach (string props in propertyArray)
                    //{
                    //    userListNode.Value.TryGetValue(props, out userProperty);
                    //    user.Properties[props].Value = userProperty;
                    //}
                    //user.CommitChanges();


                    newUser.CommitChanges();
                    newUser.Invoke("SetPassword", new object[] { userPassword });
                    newUser.CommitChanges();
                    int val = (int)newUser.Properties["userAccountControl"].Value;
                    newUser.Properties["userAccountControl"].Value = val | 0x0200;
                    newUser.CommitChanges();
                    de.Close();
                    newUser.Close();
                    return true;
                }
                catch (Exception ex)
                {
                    string err = ex.Message.ToString();

                    // MessageBox.Show("Velde.Utilities.AD.createUserAccount():\n\n" + err);

                    return false;
                }
                /*
                    //Add this to the create account method
                    int val = (int)newUser.Properties["userAccountControl"].Value; 
                         //newUser is DirectoryEntry object
                    newUser.Properties["userAccountControl"].Value = val | 0x80000; 
                        //ADS_UF_TRUSTED_FOR_DELEGATION
                 
                 * 
                 * UserAccountControlFlags
                 * CONST   HEX
                    -------------------------------
                    SCRIPT 0x0001
                    ACCOUNTDISABLE 0x0002
                    HOMEDIR_REQUIRED 0x0008
                    LOCKOUT 0x0010
                    PASSWD_NOTREQD 0x0020
                    PASSWD_CANT_CHANGE 0x0040
                    ENCRYPTED_TEXT_PWD_ALLOWED 0x0080
                    TEMP_DUPLICATE_ACCOUNT 0x0100
                    NORMAL_ACCOUNT 0x0200
                    INTERDOMAIN_TRUST_ACCOUNT 0x0800
                    WORKSTATION_TRUST_ACCOUNT 0x1000
                    SERVER_TRUST_ACCOUNT 0x2000
                    DONT_EXPIRE_PASSWORD 0x10000
                    MNS_LOGON_ACCOUNT 0x20000
                    SMARTCARD_REQUIRED 0x40000
                    TRUSTED_FOR_DELEGATION 0x80000
                    NOT_DELEGATED 0x100000
                    USE_DES_KEY_ONLY 0x200000
                    DONT_REQ_PREAUTH 0x400000
                    PASSWORD_EXPIRED 0x800000
                    TRUSTED_TO_AUTH_FOR_DELEGATION 0x1000000
                 * */
            }
            public bool CreateUserAccount(string ouPath, Dictionary<string, string> properties)
            {
                // properties contians key pairs with the key matching a field within Active driectory and the value, the value to be inserted


                try
                {
                    if (!DirectoryEntry.Exists("LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath))
                    {

                        DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                        DirectoryEntry newUser = entry.Children.Add("CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A"), "user");
                        foreach (KeyValuePair<string, string> kvp in properties)
                        {
                            newUser.Properties[kvp.Key.ToString()].Value = System.Web.HttpUtility.UrlEncode(kvp.Value.ToString()).Replace("+", " ").Replace("*", "%2A");
                        }
                        newUser.CommitChanges();
                        newUser.Invoke("SetPassword", new object[] { System.Web.HttpUtility.UrlEncode(properties["password"].ToString()).Replace("+", " ").Replace("*", "%2A") });
                        newUser.CommitChanges();
                        int val = (int)newUser.Properties["userAccountControl"].Value;
                        newUser.Properties["userAccountControl"].Value = val | 0x0200;
                        newUser.CommitChanges();
                        entry.Close();
                        newUser.Close();
                        return true;
                    }
                    else
                    {
                        MessageBox.Show("CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + " group already exists from adding");
                        return false;
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message.ToString() + "issue create user LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath);
                    return false;
                }

                /*
                    //Add this to the create account method
                    int val = (int)newUser.Properties["userAccountControl"].Value; 
                         //newUser is DirectoryEntry object
                    newUser.Properties["userAccountControl"].Value = val | 0x80000; 
                        //ADS_UF_TRUSTED_FOR_DELEGATION
                 
                 * 
                 * UserAccountControlFlags
                 * CONST   HEX
                    -------------------------------
                    SCRIPT 0x0001
                    ACCOUNTDISABLE 0x0002
                    HOMEDIR_REQUIRED 0x0008
                    LOCKOUT 0x0010
                    PASSWD_NOTREQD 0x0020
                    PASSWD_CANT_CHANGE 0x0040
                    ENCRYPTED_TEXT_PWD_ALLOWED 0x0080
                    TEMP_DUPLICATE_ACCOUNT 0x0100
                    NORMAL_ACCOUNT 0x0200
                    INTERDOMAIN_TRUST_ACCOUNT 0x0800
                    WORKSTATION_TRUST_ACCOUNT 0x1000
                    SERVER_TRUST_ACCOUNT 0x2000
                    DONT_EXPIRE_PASSWORD 0x10000
                    MNS_LOGON_ACCOUNT 0x20000
                    SMARTCARD_REQUIRED 0x40000
                    TRUSTED_FOR_DELEGATION 0x80000
                    NOT_DELEGATED 0x100000
                    USE_DES_KEY_ONLY 0x200000
                    DONT_REQ_PREAUTH 0x400000
                    PASSWORD_EXPIRED 0x800000
                    TRUSTED_TO_AUTH_FOR_DELEGATION 0x1000000
                 * */
            }
        }

        public class ObjectADSqlsyncGroup
        {
            public void ExecuteGroupSync(GroupSynch groupsyn, ToolSet tools, LogFile log, Form1 gui)
            {
                string debug = "";
                SqlDataReader debugreader;
                ArrayList debuglist = new ArrayList();
                int debugfieldcount;
                string debugrecourdcount;
                int i;
                StopWatch time = new StopWatch();
                SqlCommand sqldebugComm;

                string groupapp = groupsyn.Group_Append;
                string groupOU = groupsyn.BaseGroupOU;
                string sAMAccountName = "";
                string description = "";
                string sqlgroupsTable = "#FHC_SQLgroupsTable";
                string adGroupsTable = "#FHC_ADgroupsTable";
                string dc = groupOU.Substring(groupOU.IndexOf("DC"));
                string groupDN;
                string groupsTable;
                SqlDataReader add;
                SqlDataReader delete;
                SqlDataReader update;
                ArrayList adUpdateKeys = new ArrayList();
                ArrayList sqlUpdateKeys = new ArrayList();
                ArrayList fields = new ArrayList();
                DataTable groupsDataTable = new DataTable();
                Dictionary<string, string> groupObject = new Dictionary<string, string>();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + groupsyn.DataServer + ";Initial Catalog=" + groupsyn.DBCatalog + ";Integrated Security=SSPI;");



                sqlConn.Open();
                // Setup the OU for the program
                tools.CreateOURecursive("OU=" + groupapp + "," + groupOU, log);

                // grab list of groups from SQL insert into a temp table
                SqlCommand sqlComm = new SqlCommand();
                if (groupsyn.Group_where == "")
                {
                    sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + groupsyn.Group_sAMAccount + ") AS " + groupsyn.Group_sAMAccount + ", RTRIM(" + groupsyn.Group_CN + ") + '" + groupapp + "' AS " + groupsyn.Group_CN + " INTO " + sqlgroupsTable + " FROM " + groupsyn.Group_dbTable + " ORDER BY " + groupsyn.Group_CN, sqlConn);
                }
                else
                {
                    sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + groupsyn.Group_sAMAccount + ") AS " + groupsyn.Group_sAMAccount + ", RTRIM(" + groupsyn.Group_CN + ") + '" + groupapp + "' AS " + groupsyn.Group_CN + " INTO " + sqlgroupsTable + " FROM " + groupsyn.Group_dbTable + " WHERE " + groupsyn.Group_where + " ORDER BY " + groupsyn.Group_CN, sqlConn);
                }

                sqlComm.ExecuteNonQuery();


                //sqldebugComm = new SqlCommand("select count(" + groupsyn.Group_sAMAccount + ") FROM " + sqlgroupsTable, sqlConn);
                //debugreader = sqldebugComm.ExecuteReader();
                //debugfieldcount = debugreader.FieldCount;
                //while (debugreader.Read())
                //{
                //    for (i = 0; i < debugfieldcount; i++)
                //    {
                //        debug += (string)debugreader[0].ToString();
                //    }
                //}
                //MessageBox.Show(debug);
                //debugreader.Close();


                // generate a list of fields to ask from AD
                adUpdateKeys.Add("description");
                adUpdateKeys.Add("CN");


                // grab groups from AD
                //time.Start();
                groupsDataTable = tools.EnumerateGroupsInOUDataTable("OU=" + groupapp + "," + groupOU, adUpdateKeys, adGroupsTable);
                //time.Stop();
                gui.Refresh();
                //MessageBox.Show("got " + groupsLinkedList.Count + "groups from ou in " + time.GetElapsedTime());
                // insert groups from AD into a temp table
                if (groupsDataTable.Rows.Count > 0)
                {
                    //time.Start();
                    groupsTable = tools.Create_Table(groupsDataTable, adGroupsTable, sqlConn);
                    //time.Stop();
                    //MessageBox.Show("temp table loaded " + groupsLinkedList.Count + " in " + time.GetElapsedTime());


                    //debug = " groups table  data import from AD \n";
                    //sqldebugComm = new SqlCommand("select top 20 * FROM " + groupsTable, sqlConn);
                    //debugreader = sqldebugComm.ExecuteReader();
                    //debugfieldcount = debugreader.FieldCount;
                    ////debugrecourdcount = debugreader.RecordsAffected.ToString();
                    //for (i = 0; i < debugfieldcount; i++)
                    //{
                    //    debug += debugreader.GetName(i);
                    //}
                    //debug += "\n";
                    //while (debugreader.Read())
                    //{
                    //    for (i = 0; i < debugfieldcount; i++)
                    //    {
                    //        debug += (string)debugreader[i] + ",";
                    //    }
                    //    debug += "\n";
                    //}
                    //sqldebugComm = new SqlCommand("select count(" + debugreader.GetName(1) + ") FROM " + groupsTable, sqlConn);
                    //debugreader.Close();
                    //debugrecourdcount = sqldebugComm.ExecuteScalar().ToString(); 
                    //MessageBox.Show("table " + groupsTable + " has " + debugrecourdcount + " records \n " + debugfieldcount + " fields \n sample data" + debug);



                    //debug = " groups from SQL to compare against AD \n";
                    //sqldebugComm = new SqlCommand("select top 20 * FROM " + sqlgroupsTable, sqlConn);
                    //debugreader = sqldebugComm.ExecuteReader();
                    //debugfieldcount = debugreader.FieldCount;
                    ////debugrecourdcount = debugreader.RecordsAffected.ToString();
                    //for (i = 0; i < debugfieldcount; i++)
                    //{
                    //    debug += debugreader.GetName(i);
                    //}
                    //debug += "\n";
                    //while (debugreader.Read())
                    //{
                    //    for (i = 0; i < debugfieldcount; i++)
                    //    {
                    //        debug += (string)debugreader[i] + ",";
                    //    }
                    //    debug += "\n";
                    //}
                    //sqldebugComm = new SqlCommand("select count(" + debugreader.GetName(1) + ") FROM " + sqlgroupsTable, sqlConn);
                    //debugreader.Close();
                    //debugrecourdcount = sqldebugComm.ExecuteScalar().ToString();
                    //MessageBox.Show("table " + sqlgroupsTable + " has " + debugrecourdcount + " records \n " + debugfieldcount + " fields \n sample data" + debug);


                    // does not get columns from a temp table as they are not in the system objects database
                    //debuglist = tools.GetColumns(groupsyn.DataServer, groupsyn.DBCatalog, sqlgroupsTable, sqlConn);
                    //debug = " columns \n";
                    //foreach (string a in debuglist)
                    //{
                    //    debug += a + "\n";
                    //}
                    //MessageBox.Show(debug);



                    //time.Start();
                    add = tools.QueryNotExists(sqlgroupsTable, groupsTable, sqlConn, groupsyn.Group_CN, adUpdateKeys[1].ToString());

                    //time.Stop();
                    //MessageBox.Show("add query" + time.GetElapsedTime());



                    //debug = "cols to add \n";
                    //while (add.Read())
                    //{
                    //    debug += (string)add[0] + "\n";
                    //}
                    //MessageBox.Show(debug);

                    // add nodes to AD

                    //time.Start();
                   // i = 0;
                    while (add.Read())
                    {
                        //i++;
                        sAMAccountName = (string)add[1].ToString().Trim();
                        description = (string)add[0].ToString().Trim();
                        groupObject.Add("sAMAccountName", sAMAccountName);
                        groupObject.Add("CN", sAMAccountName);
                        groupObject.Add("description", description);
                        tools.CreateGroup("OU=" + groupapp + "," + groupOU, groupObject, log);
                        // log.transactions.Add("Group added ;" + sAMAccountName + ",OU=" + groupapp + "," + groupOU + ";" + description);
                        //if (i % 1000 == 0)
                        //{
                        //    // FORGET the real progress bar for now groupsyn.progress = i;
                        //    gui.group_result1.Text = "Adding cause im still ALIVE !!!" + i;
                        //    gui.Refresh();
                        //    //MessageBox.Show("adding now at item " + i);
                        //}
                        groupObject.Clear();
                    }
                    //time.Stop();
                    //MessageBox.Show("add " + i + " objects " + time.GetElapsedTime());
                    add.Close();


                    //time.Start();
                    delete = tools.QueryNotExists(groupsTable, sqlgroupsTable, sqlConn, adUpdateKeys[1].ToString(), groupsyn.Group_CN);
                    // delete groups in AD
                   // i = 0;
                    while (delete.Read())
                    {
                       // i++;
                        tools.DeleteGroup("OU=" + groupapp + "," + groupOU, (string)delete[adUpdateKeys[1].ToString()].ToString().Trim(), log);
                        // log.transactions.Add("Group deleted ;" + (string)delete[adUpdateKeys[1].ToString()].ToString().Trim() + ",OU=" + groupapp + groupOU);
                        //if (i % 1000 == 0)
                        //{
                        //    // FORGET the real progress bar for now groupsyn.progress = i;
                        //    gui.group_result1.Text = "Deleting cause im still ALIVE !!!" + i;
                        //    gui.Refresh();
                        //    //MessageBox.Show("Deleting now at item " + i);
                        //}
                    }
                    delete.Close();
                    //time.Stop();
                    //MessageBox.Show("Delete " + i + " objects " + time.GetElapsedTime());


                    // Get columns from sqlgroupsTable temp table in database get columns deprcated in favor of manual building due to cannot figure out how to get the columns of a temporary table
                    // SQLupdateKeys = tools.GetColumns(groupsyn.DataServer, groupsyn.DBCatalog, sqlgroupsTable);
                    // make the list of fields for the sql to check when updating
                    sqlUpdateKeys.Add(groupsyn.Group_sAMAccount);
                    sqlUpdateKeys.Add(groupsyn.Group_CN);
                    //time.Start();
                    // update assumes the both ADupdateKeys and SQLupdateKeys have the same fields, listed in the same order check  call to EnumerateGroupsInOU if this is wrong should be sAMAccountName, CN matching the SQL order
                    update = tools.CheckUpdate(sqlgroupsTable, groupsTable, groupsyn.Group_CN, adUpdateKeys[1].ToString(), sqlUpdateKeys, adUpdateKeys, sqlConn);
                    //time.Stop();
                    //MessageBox.Show("update query" + time.GetElapsedTime());

                    //int  j = 0;
                    //debug = "Records to Update ";
                    //while (update.Read() && j < 30)
                    //{
                    //    j++;
                    //    for (i = 0; i < update.FieldCount; i++)
                    //    {
                    //        debug += (string)update[i].ToString() + ",";
                    //    }
                    //    debug += "\n";
                    //}
                    //debug += update.RecordsAffected;
                    //MessageBox.Show(debug);



                    // update groups in ad
                    //time.Start();
                   // i = 0;
                    // last record which matches the primary key is the one which gets inserted into the database
                    while (update.Read())
                    {
                        // any duplicate records will attempt to be updated if slow runtimes are a problem this might be an issue
                       // i++;
                        sAMAccountName = (string)update[1].ToString().Trim();
                        description = (string)update[0].ToString().Trim();
                        groupObject.Add("sAMAccountName", sAMAccountName);
                        groupObject.Add("CN", sAMAccountName);
                        groupObject.Add("description", description);

                        if (tools.Exists("CN=" + groupObject["CN"] + ", OU=" + groupapp + "," + groupOU) == true)
                        {
                            // group exists in place just needs updating
                            tools.UpdateGroup("OU=" + groupapp + "," + groupOU, groupObject, log);
                            // log.transactions.Add("Group update ; " + sAMAccountName + ",OU=" + groupapp + "," + groupOU + ";" + description);
                        }
                        else
                        {
                            // find it its on the server somewhere we will log the exception
                            groupDN = tools.GetObjectDistinguishedName(objectClass.group, returnType.distinguishedName, groupObject["CN"], dc);
                            // what if user is disabled will user mapping handle it?
                            // groups needs to be moved and updated
                            // tools.MoveADObject(groupDN, "LDAP://OU=" + groupapp + ',' + groupOU);
                            // tools.UpdateGroup("OU=" + groupapp + "," + groupOU, groupObject);
                            log.errors.Add("Group cannot be updated user probabally should be in ; " + "OU=" + groupapp + "," + groupOU + " ; but was found in ; " + groupDN);
                        }
                        //if (i % 1000 == 0)
                        //{
                        //    // FORGET the real progress bar for now groupsyn.progress = i;
                        //    gui.group_result1.Text = "updating cause im still ALIVE !!!" + i;
                        //    gui.Refresh();
                        //    //MessageBox.Show("updating now at item " + i);
                        //}
                        groupObject.Clear();
                    }
                    update.Close();
                    //time.Stop();
                    //MessageBox.Show("update objects somehow found " + i + " objects to finished in "  + time.GetElapsedTime());
                }
                else
                {
                    sqlComm = new SqlCommand("select * FROM " + sqlgroupsTable, sqlConn);
                    add = sqlComm.ExecuteReader();
                    //time.Start();
                    // i = 0;
                    while (add.Read())
                    {
                        //i++;
                        groupObject.Add("sAMAccountName", (string)add[1]);
                        groupObject.Add("CN", (string)add[1]);
                        groupObject.Add("description", (string)add[0]);
                        tools.CreateGroup("OU=" + groupapp + "," + groupOU, groupObject, log);
                        // log.transactions.Add("Group added ;" + groupObject["sAMAccountName"] + ",OU=" + groupapp + "," + groupOU + ";" + groupObject["description"]);

                        groupObject.Clear();
                        //if (i % 500 == 0)
                        //{
                        //    // FORGET the real progress bar for now groupsyn.progress = i;
                        //    gui.group_result1.AppendText("add cause im still ALIVE !!!" + i);
                        //    gui.Refresh();
                        //    // MessageBox.Show("avoiding message pumping add progress now at item " + i);
                        //}
                    }
                    //time.Stop();
                    //MessageBox.Show("initial add objects " + i + " time taken" + time.GetElapsedTime());
                    add.Close();
                }

                // users section
                string sqlgroupMembersTable = "#sqlusersTable";
                string ADgroupMembersTable = "#ADusersTable";
                SqlDataReader sqlgroups;
                DataTable ADusers = new DataTable();

                // grab users data from sql
                if (groupsyn.User_where == "")
                {
                    sqlComm = new SqlCommand("SELECT DISTINCT 'CN=' + RTRIM(" + groupsyn.User_sAMAccount + ") + '," + groupsyn.BaseUserOU + "' AS " + groupsyn.User_sAMAccount + ", RTRIM(" + groupsyn.User_Group_Reference + ") + '" + groupapp + "' AS " + groupsyn.User_Group_Reference + " INTO " + sqlgroupMembersTable + " FROM " + groupsyn.User_dbTable, sqlConn);
                }
                else
                {
                    sqlComm = new SqlCommand("SELECT DISTINCT 'CN=' + RTRIM(" + groupsyn.User_sAMAccount + ") + '," + groupsyn.BaseUserOU + "' AS " + groupsyn.User_sAMAccount + ", RTRIM(" + groupsyn.User_Group_Reference + ") + '" + groupapp + "' AS " + groupsyn.User_Group_Reference + " INTO " + sqlgroupMembersTable + " FROM " + groupsyn.User_dbTable + " WHERE " + groupsyn.User_where, sqlConn);
                }
                sqlComm.ExecuteNonQuery();

                // populate datatable with users from AD groups by looping thru the list of groups from SQL and loading the cross referenced AD group members
                sqlComm = new SqlCommand("select " + groupsyn.Group_CN + " FROM " + sqlgroupsTable, sqlConn);
                sqlgroups = sqlComm.ExecuteReader();
                while (sqlgroups.Read())
                {
                       // hopefully merge acts as an append
                    ADusers.Merge(tools.EnumerateUsersInGroupDataTable((string)sqlgroups[0], ",OU=" + groupapp + "," + groupOU, groupsyn.User_sAMAccount, groupsyn.User_Group_Reference, ADgroupMembersTable));
                }
                sqlgroups.Close();

                // make the temp table for ou comparisons
                tools.Create_Table(ADusers, ADgroupMembersTable, sqlConn);



                debug = " total users in groups from SQL \n";
                sqldebugComm = new SqlCommand("select top 20 * FROM " + sqlgroupMembersTable, sqlConn);
                debugreader = sqldebugComm.ExecuteReader();
                debugfieldcount = debugreader.FieldCount;
                for (i = 0; i < debugfieldcount; i++)
                {
                    debug += debugreader.GetName(i);
                }
                debug += "\n";
                while (debugreader.Read())
                {
                    for (i = 0; i < debugfieldcount; i++)
                    {
                        debug += (string)debugreader[i] + ",";
                    }
                    debug += "\n";
                }
                sqldebugComm = new SqlCommand("select count(" + ADusers.Columns[0].ColumnName + ") FROM " + sqlgroupMembersTable, sqlConn);
                debugreader.Close();
                debugrecourdcount = sqldebugComm.ExecuteScalar().ToString();
                MessageBox.Show("table " + sqlgroupMembersTable + " has " + debugrecourdcount + " records \n " + debugfieldcount + " fields \n sample data" + debug);

                debug = " total users in groups from AD \n";
                sqldebugComm = new SqlCommand("select top 20 * FROM " + ADgroupMembersTable, sqlConn);
                debugreader = sqldebugComm.ExecuteReader();
                debugfieldcount = debugreader.FieldCount;
                for (i = 0; i < debugfieldcount; i++)
                {
                    debug += debugreader.GetName(i);
                }
                debug += "\n";
                while (debugreader.Read())
                {
                    for (i = 0; i < debugfieldcount; i++)
                    {
                        debug += (string)debugreader[i] + ",";
                    }
                    debug += "\n";
                }
                sqldebugComm = new SqlCommand("select count(" + ADusers.Columns[0].ColumnName + ") FROM " + ADgroupMembersTable, sqlConn);
                debugreader.Close();
                debugrecourdcount = sqldebugComm.ExecuteScalar().ToString();
                MessageBox.Show("table " + ADgroupMembersTable + " has " + debugrecourdcount + " records \n " + debugfieldcount + " fields \n sample data" + debug);


                
                // compare and add/remove
                add = tools.QueryNotExists(sqlgroupMembersTable, ADgroupMembersTable, sqlConn, groupsyn.User_Group_Reference, ADusers.Columns[1].ColumnName);
                while (add.Read())
                {
                    tools.AddUserToGroup((string)add[0], "CN=" + (string)add[1] + ",OU=" + groupapp + "," + groupOU, log);
                    // log.transactions.Add("User added ;" + (string)add[0] + ",OU=" + groupapp + "," + groupOU + ";" + (string)add[1]);
                    groupObject.Clear();
                }
                add.Close();

                delete = tools.QueryNotExists(ADgroupMembersTable, sqlgroupMembersTable, sqlConn, ADusers.Columns[1].ColumnName, groupsyn.User_Group_Reference);
                // delete groups in AD
                while (delete.Read())
                {

                    tools.RemoveUserFromGroup((string)delete[0], (string)delete[1], log);
                    // log.transactions.Add("User removed ;" + (string)delete[adUpdateKeys[1].ToString()].ToString().Trim() + ",OU=" + groupapp + groupOU);

                }
                delete.Close();
                sqlConn.Close();
            }
            public void ExecuteUserSync(UserSynch usersyn, ToolSet tools, LogFile log, Form1 gui)
            {
                string debug = "";
                SqlDataReader debugReader;
                ArrayList debugList = new ArrayList();
                int debugFieldCount;
                string debugRecordCount;
                int i;
                StopWatch time = new StopWatch();
                SqlCommand sqlDebugComm;

                string baseOU = usersyn.BaseUserOU;
                string DC = baseOU.Substring(baseOU.IndexOf("DC"));
                string sqlForCustomFields = "";

                SqlDataReader add;
                SqlDataReader delete;
                SqlDataReader update;

                ArrayList adUpdateKeys = new ArrayList();
                ArrayList sqlUpdateKeys = new ArrayList();
                ArrayList fields = new ArrayList();
                Dictionary<string, string> userObject = new Dictionary<string, string>();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + usersyn.DataServer + ";Initial Catalog=" + usersyn.DBCatalog + ";Integrated Security=SSPI;");
                

                string sqlUsersTable = "#sqlusersTable";
                string adUsersTable = "#ADusersTable";
                //SqlDataReader sqlusers;
                SqlCommand sqlComm;
                DataTable adUsers = new DataTable();
                ArrayList userProperties = new ArrayList();

                sqlConn.Open();
                tools.CreateOURecursive(usersyn.BaseUserOU, log);
                tools.CreateOURecursive(usersyn.UserHoldingTank, log);
                userObject.Add("sAMAccountName", usersyn.UniversalGroup.Remove(0,3).Remove(usersyn.UniversalGroup.IndexOf(",") -3));
                userObject.Add("CN", usersyn.UniversalGroup.Remove(0,3).Remove(usersyn.UniversalGroup.IndexOf(",") - 3));
                userObject.Add("description", "Universal Group For Users");
                // creates the group if it does not exist
                tools.CreateGroup(usersyn.UniversalGroup.Remove(0,usersyn.UniversalGroup.IndexOf(",") + 1), userObject, log);

                
                // need to add this field first to use as a primary key when checking for existance in AD
                userProperties.Add("sAMAccountName");
                for (i = 0; i < usersyn.UserCustoms.Rows.Count; i++)
                {
                    // build fields to pull back from ad
                    userProperties.Add(usersyn.UserCustoms.Rows[i][0].ToString());
                    //create props from rows in usercustoms datatable our column names match the appropriate fields in AD and SQL
                    if (usersyn.UserCustoms.Rows[i][1].ToString() != "Static Value")
                    {
                        sqlForCustomFields += ", RTRIM(" + usersyn.UserCustoms.Rows[i][1].ToString() + ") AS " + usersyn.UserCustoms.Rows[i][0].ToString();
                    }
                    // static fields get static values for the table to get updated
                    else
                    {
                        sqlForCustomFields += ", '" + usersyn.UserCustoms.Rows[i][2].ToString() + "' AS " + usersyn.UserCustoms.Rows[i][0].ToString();
                    }
                }

                // grab users data from sql
                if (usersyn.User_where == "")
                {
                    sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + usersyn.User_sAMAccount + ") AS sAMAccountName" +
                        ", RTRIM(" + usersyn.User_sAMAccount + ") AS CN" +
                        ", RTRIM(" + usersyn.User_Lname + ") AS sn" +
                        ", RTRIM(" + usersyn.User_Fname + ") AS givenname" +
                        ", RTRIM(" + usersyn.User_Mobile + ") AS homephone" +
                        ", RTRIM(" + usersyn.User_State + ") AS st"+
                        ", RTRIM(" + usersyn.User_Address + ") AS streetaddress"+
                        ", RTRIM(" + usersyn.User_city + ") AS l" +
                        ", RTRIM(" + usersyn.User_Zip + ") AS postalcode" +
                        ", RTRIM(" + usersyn.User_password + ") AS password" + 
                        sqlForCustomFields +
                        " INTO " + sqlUsersTable + " FROM " + usersyn.User_dbTable, sqlConn);
                }
                else
                {
                    sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + usersyn.User_sAMAccount + ") AS sAMAccountName" +
                        ", RTRIM(" + usersyn.User_sAMAccount + ") AS CN" +
                        ", RTRIM(" + usersyn.User_Lname + ") AS sn" +
                        ", RTRIM(" + usersyn.User_Fname + ") AS givenname" +
                        ", RTRIM(" + usersyn.User_Mobile + ") AS homephone" +
                        ", RTRIM(" + usersyn.User_State + ") AS st"+
                        ", RTRIM(" + usersyn.User_Address + ") AS streetaddress"+
                        ", RTRIM(" + usersyn.User_city + ") AS l" +
                        ", RTRIM(" + usersyn.User_Zip + ") AS postalcode" +
                        ", RTRIM(" + usersyn.User_password + ") AS password" +
                        sqlForCustomFields +
                        " INTO " + sqlUsersTable + " FROM " + usersyn.User_dbTable +
                        " WHERE " + usersyn.User_where, sqlConn);
                }
                sqlComm.ExecuteNonQuery();

                // set up fields to pull back from AD                 
                userProperties.Add("CN");
                userProperties.Add("sn");
                userProperties.Add("givenname");
                userProperties.Add("homephone");
                userProperties.Add("st");
                userProperties.Add("streetaddress");
                userProperties.Add("l");
                userProperties.Add("postalcode");
                userProperties.Add("distinguishedName");

                adUsers = tools.EnumerateUsersInOUDataTable(usersyn.BaseUserOU, userProperties, adUsersTable);
                // make the temp table for ou comparisons
                tools.Create_Table(adUsers, adUsersTable, sqlConn);




                debug = " total users from sql \n";
                sqlDebugComm = new SqlCommand("select top 20 * FROM " + sqlUsersTable, sqlConn);
                debugReader = sqlDebugComm.ExecuteReader();
                debugFieldCount = debugReader.FieldCount;
                for (i = 0; i < debugFieldCount; i++)
                {
                    debug += debugReader.GetName(i) + ", ";
                }
                debug += "\n";
                while (debugReader.Read())
                {
                    for (i = 0; i < debugFieldCount; i++)
                    {
                        debug += (string)debugReader[i].ToString() + ", ";
                    }
                    debug += "\n";
                }
                sqlDebugComm = new SqlCommand("select count(sAMAccountName) FROM " + sqlUsersTable, sqlConn);
                debugReader.Close();
                debugRecordCount = sqlDebugComm.ExecuteScalar().ToString();
                MessageBox.Show("table " + sqlUsersTable + " has " + debugRecordCount + " records \n " + debugFieldCount + " fields \n sample data" + debug);


                debug = "";
                debug = " total users from AD \n";
                sqlDebugComm = new SqlCommand("select top 20 * FROM " + adUsersTable, sqlConn);
                debugReader = sqlDebugComm.ExecuteReader();
                debugFieldCount = debugReader.FieldCount;
                for (i = 0; i < debugFieldCount; i++)
                {
                    debug += debugReader.GetName(i) + ", ";
                }
                debug += "\n";
                while (debugReader.Read())
                {
                    for (i = 0; i < debugFieldCount; i++)
                    {
                        debug += (string)debugReader[i] + ", ";
                    }
                    debug += "\n";
                }
                sqlDebugComm = new SqlCommand("select count(" + adUsers.Columns[0].ColumnName + ") FROM " + adUsersTable, sqlConn);
                debugReader.Close();
                debugRecordCount = sqlDebugComm.ExecuteScalar().ToString();
                MessageBox.Show("table " + adUsersTable + " has " + debugRecordCount + " records \n " + debugFieldCount + " fields \n sample data" + debug);






                // compare and add/remove
                add = tools.QueryNotExists(sqlUsersTable, adUsersTable, sqlConn, "sAMAccountName", adUsers.Columns[0].ColumnName);


                debug = "Gunna Add stuff \n";
                debugFieldCount = add.FieldCount;
                for (i = 0; i < debugFieldCount; i++)
                {
                    debug += add.GetName(i) + ", ";
                }
                debug += "\n";
                int j = 0;
                while (add.Read() && j < 20)
                {
                    for (i = 0; i < debugFieldCount; i++)
                    {
                        debug += (string)add[i] + ", ";
                    }
                    debug += "\n";
                    j++;
                }

                debugReader.Close();
                MessageBox.Show("table " + adUsersTable + "\n " + debugFieldCount + " fields \n sample data" + debug);

                tools.CreateUserAccount(usersyn.UserHoldingTank, add, usersyn.UniversalGroup, usersyn, log);
                add.Close();

                delete = tools.QueryNotExists(sqlUsersTable, adUsersTable, sqlConn, usersyn.User_Lname, adUpdateKeys[1].ToString());
                // delete groups in AD
                while (delete.Read())
                { 
                    tools.DeleteUserAccount((string)delete[0], (string)delete[1], log);
                    // log.transactions.Add("User removed ;" + (string)delete[adUpdateKeys[1].ToString()].ToString().Trim()); 
                }
                delete.Close();
                sqlConn.Close();
            }
        }

        // create objects to hold save data
        GroupSynch groupconfig = new GroupSynch();
        UserSynch userconfig = new UserSynch();
        executionOrder execution = new executionOrder();
        UserStateChange usermapping = new UserStateChange();                          
        ToolSet tools = new ToolSet();
        LogFile log = new LogFile();
        ObjectADSqlsyncGroup groupSyncr = new ObjectADSqlsyncGroup();

        //private void group_git_er_done_Click(object sender, System.EventArgs e)
        //{

        //    //PURPOSE
        //    //
        //    // get all users in the right groups regardless of what OU they are in
        //    // ensure all the groups are in the right OU's
        //    //
        //    //

        //    // variables for the group setup and iteration
        //    LinkedList<Dictionary<string, string>> DBgroups = new LinkedList<Dictionary<string, string>>();
        //    Dictionary<string, string> DBgroupDictionary;
        //    LinkedListNode<Dictionary<string, string>> DBgroupsNode;
        //    string groupDN;
        //    ArrayList compare = new ArrayList();
        //    ArrayList groupProperties = new ArrayList();
        //    ArrayList listupdate = new ArrayList();
        //    LinkedList<Dictionary<string, string>> ADgroups = new LinkedList<Dictionary<string, string>>();
        //    Dictionary<string, string> ADgroupsDictionary = new Dictionary<string, string>();
        //    LinkedListNode<Dictionary<string, string>> ADgroupsNode;

        //    // variables for the user synch
        //    LinkedList<Dictionary<string, string>> DBgroupList = new LinkedList<Dictionary<string, string>>();
        //    LinkedListNode<Dictionary<string, string>> DBgroupListNode;
        //    Dictionary<string, string> UsergroupDictionary;
        //    LinkedList<Dictionary<string, string>> DBUsers = new LinkedList<Dictionary<string, string>>();
        //    LinkedList<Dictionary<string, string>> ADgroupUsers = new LinkedList<Dictionary<string, string>>();
        //    LinkedListNode<Dictionary<string, string>> ADUsersNode;
        //    LinkedListNode<Dictionary<string, string>> DBUsersNode;
        //    string DC = groupconfig.BaseGroupOU.Substring(groupconfig.BaseGroupOU.IndexOf("DC"));

        //    // Setup the OU for the program
        //    tools.createOURecursive("OU=" + groupconfig.Group_Append + "," + groupconfig.BaseGroupOU);

        //    // grab list of groups from SQL
        //    SqlConnection sqlConn = new SqlConnection("Data Source=" + groupconfig.DataServer.ToString() + ";Initial Catalog=" + groupconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");

        //    sqlConn.Open();
        //    // create the command object
        //    SqlCommand sqlComm = new SqlCommand();
        //    if (groupconfig.Group_where == "")
        //    {
        //        sqlComm = new SqlCommand("SELECT " + groupconfig.Group_sAMAccount + ", " + groupconfig.Group_CN + " FROM " + groupconfig.Group_dbTable, sqlConn);
        //    }
        //    else
        //    {
        //        sqlComm = new SqlCommand("SELECT " + groupconfig.Group_sAMAccount + ", " + groupconfig.Group_CN + " FROM " + groupconfig.Group_dbTable + " WHERE " + groupconfig.Group_where, sqlConn);
        //    }
        //    SqlDataReader r = sqlComm.ExecuteReader();

        //    // interate thru a recordset based on query generated from text and generate the linked list of dictionary for diff
        //    while (r.Read())
        //    {
        //        DBgroupDictionary = new Dictionary<string, string>();
        //        DBgroupDictionary.Add("sAMAccountName", (string)r[groupconfig.Group_CN].ToString().Trim() + groupconfig.Group_Append);
        //        DBgroupDictionary.Add("CN", (string)r[groupconfig.Group_CN].ToString().Trim() + groupconfig.Group_Append);
        //        DBgroupDictionary.Add("description", (string)r[groupconfig.Group_sAMAccount].ToString().Trim());
        //        DBgroups.AddLast(DBgroupDictionary);

        //    }
        //    r.Close();
        //    sqlConn.Close();

        //    DBgroupsNode = DBgroups.First;
        //    while (DBgroupsNode != null)
        //    {
        //        DBgroupList.AddFirst(DBgroupsNode.Value);
        //        DBgroupsNode = DBgroupsNode.Next;
        //    }

        //    // build a list of all data gathered from the SQL command so if any field changes we wil be able to detect it in our diff
        //    DBgroupDictionary = DBgroups.First.Value;
        //    foreach (KeyValuePair<string, string> kvp in DBgroupDictionary)
        //    {
        //        groupProperties.Add(kvp.Key.ToString());
        //    }

        //    // list of keys must be fields pulled in SQL query and Group pull
        //    compare.Add("CN");
        //    compare.Add("sAMAccountName");
        //    // list of fields to synch must be pulled in SQL query and Group pull
        //    listupdate.Add("description");


        //    // grab list of groups from AD EnumerateGroupsInOU(string groupDN)
        //    ADgroups = tools.EnumerateGroupsInOU("OU=" + groupconfig.Group_Append + "," + groupconfig.BaseGroupOU, groupProperties);



        //    // diff groups
        //    tools.Diff(DBgroups, ADgroups, compare, compare, listupdate, listupdate);

        //    // Delete rogue nodes from AD
        //    ADgroupsNode = ADgroups.First;
        //    while (ADgroupsNode != null)
        //    {
        //        tools.DeleteGroup("OU=" + groupconfig.Group_Append + groupconfig.BaseGroupOU, ADgroupsNode.Value["CN"]);
        //        ADgroupsNode = ADgroupsNode.Next;
        //    }

        //    // These groups do not exist in the right place, Find them and move them or make them
        //    DBgroupsNode = DBgroups.First;
        //    while (DBgroupsNode != null)
        //    {
        //        // if it exists in place its got to get updated
        //        if (tools.Exists("CN=" + DBgroupsNode.Value["CN"] + ", OU=" + groupconfig.Group_Append + "," + groupconfig.BaseGroupOU) == true)
        //        {
        //            // update the group information it has changed
        //            tools.UpdateGroup("OU=" + groupconfig.Group_Append + "," + groupconfig.BaseGroupOU, DBgroupsNode.Value);
        //        }
        //        else
        //        {
        //            // it might be lost, go acquire their DN if they exist on the server
        //            groupDN = tools.GetObjectDistinguishedName(objectClass.group, returnType.distinguishedName, DBgroupsNode.Value[compare[1].ToString()], DC);
        //            if (groupDN != string.Empty)
        //            {
        //                // groups exists move it to the correct spot
        //                tools.MoveADObject(groupDN, "LDAP://OU=" + groupconfig.Group_Append + ',' + groupconfig.BaseGroupOU);
        //                // group may also have the wrong information
        //                tools.UpdateGroup("OU=" + groupconfig.Group_Append + "," + groupconfig.BaseGroupOU, DBgroupsNode.Value);


        //            }
        //            else
        //            {
        //                // groups really doos not exist create it
        //                // CreateGroup("OU=" + groupconfig.Group_Append + "," + groupconfig.BaseGroupOU, DBgroupsNode.Value[compare[1].ToString()], DBgroupsNode.Value[compare[0].ToString()]);
        //                tools.CreateGroup("OU=" + groupconfig.Group_Append + "," + groupconfig.BaseGroupOU, DBgroupsNode.Value);
        //            }
        //        }
        //        DBgroupsNode = DBgroupsNode.Next;
        //    }


        //    // show final output for groups
        //    group_result1.AppendText("add these groups to AD \n");
        //    DBgroupsNode = DBgroups.First;
        //    while (DBgroupsNode != null)
        //    {
        //        group_result1.AppendText(DBgroupsNode.Value[compare[0].ToString()] + "\n");
        //        DBgroupsNode = DBgroupsNode.Next;
        //    }

        //    ADgroupsNode = ADgroups.First;
        //    group_result2.AppendText("delete these groups from AD \n");
        //    while (ADgroupsNode != null)
        //    {
        //        group_result2.AppendText(ADgroupsNode.Value[compare[0].ToString()] + "\n");
        //        ADgroupsNode = ADgroupsNode.Next;
        //    }


        //    // list of groups from SQL DBgroupList retained from above

        //    // grab list of users from SQL for the group set
        //    sqlConn = new SqlConnection("Data Source=" + groupconfig.DataServer.ToString() + ";Initial Catalog=" + groupconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");

        //    sqlConn.Open();
        //    // create the command object
        //    sqlComm = new SqlCommand();
        //    if (groupconfig.User_where == "")
        //    {
        //        sqlComm = new SqlCommand("SELECT " + groupconfig.User_sAMAccount + ", " + groupconfig.User_Lname + " FROM " + groupconfig.User_dbTable, sqlConn);
        //    }
        //    else
        //    {
        //        sqlComm = new SqlCommand("SELECT " + groupconfig.User_sAMAccount + ", " + groupconfig.User_Lname + " FROM " + groupconfig.User_dbTable + " WHERE " + groupconfig.User_where, sqlConn);
        //    }
        //    r = sqlComm.ExecuteReader();

        //    // interate thru a recordset based on query generated from text and generate the linked list of dictionary for diff
        //    while (r.Read())
        //    {
        //        UsergroupDictionary = new Dictionary<string, string>();
        //        //  groupconfig.User_Lname holds the value for the cross refernce aginst the group CN 
        //        UsergroupDictionary.Add("sAMAccountName", (string)r[groupconfig.User_sAMAccount].ToString().Trim());
        //        UsergroupDictionary.Add("CN", (string)r[groupconfig.User_Lname].ToString().Trim() + groupconfig.Group_Append);
        //        DBUsers.AddLast(UsergroupDictionary);

        //    }
        //    r.Close();
        //    sqlConn.Close();

        //    //generate a list of users for all groups in base ou
        //    DBgroupListNode = DBgroupList.First;
        //    while (DBgroupListNode != null)
        //    {
        //        // grab list of users in AD for group[x] EnumerateUsersInGroup(string ouDN) get DN o removing them will be easier
        //        ADgroupUsers = tools.linkedlistadd(ADgroupUsers, tools.EnumerateUsersInGroup("CN=" + DBgroupListNode.Value["CN"] + ",OU=" + groupconfig.Group_Append + "," + groupconfig.BaseGroupOU));
        //        DBgroupListNode = DBgroupListNode.Next;
        //    }
        //    compare.Clear();
        //    compare.Add("sAMAccountName");
        //    compare.Add("CN");

        //    // no need to check fro updates
        //    tools.Diff(DBUsers, ADgroupUsers, compare, compare);
        //    // diff users
        //    // SQL vs group[x]
        //    // add or delete update group memberships
        //    // Delete rogue nodes from AD
        //    ADUsersNode = ADgroupUsers.First;
        //    while (ADUsersNode != null)
        //    {
        //        // we have their Distinguished Name so we can send it righ off to the remove
        //        tools.RemoveUserFromGroup(ADUsersNode.Value["sAMAccountName"], ADUsersNode.Value["CN"]);
        //        ADUsersNode = ADUsersNode.Next;
        //    }
        //    DBUsersNode = DBUsers.First;
        //    while (DBUsersNode != null)
        //    {
        //        // we need to get the Distinguished name before we can add them the DBn info does not provide the DN fro where the user is
        //        tools.AddUserToGroup(tools.GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, DBUsersNode.Value["sAMAccountName"], DC), DBUsersNode.Value["CN"]);
        //        DBUsersNode = DBUsersNode.Next;
        //    }
        //}


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
        private void users_holdingTank_TextChanged(object sender, EventArgs e)
        {
            if (users_holdingTank.Text.ToString() != "")
            {
                if (tools.Exists(users_holdingTank.Text.ToString()))
                {
                    userconfig.UserHoldingTank = users_holdingTank.Text.ToString();
                }
                else
                {

                    DialogResult button = MessageBox.Show("OU LDAP://" + users_baseUserOU.Text.ToString() + " does Not exist shall I create it", "Nonexistent OU", MessageBoxButtons.YesNo);

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
        
        private void users_see_test_results_Click(object sender, EventArgs e)
        {
            int i;
            StopWatch timer = new StopWatch();
            timer.Start();
            groupSyncr.ExecuteUserSync(userconfig, tools, log, this);
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
        private void users_cancel_Click(object sender, EventArgs e)
        {

            DataSet bob = new DataSet();
            DataTable mike = new DataTable();
            BindingSource bs = new BindingSource();
            bs.DataSource = mike;

            bs.DataSource = mappinggrid.DataSource;
            // mappinggrid.DataSource = bs;
            // bs.EndEdit();

            StringBuilder tableout = new StringBuilder();
            int i = 0;
            int j = 0;
            for (i = 0; i < mike.Rows.Count; i++)
            {
                for (j = 0; j < mike.Columns.Count; j++)
                {
                    tableout.Append(mike.Rows[i][j] + ", ");
                }
                tableout.Append("\n");
            }
            MessageBox.Show(tableout.ToString()); 


            for (i = 0; i < mappinggrid.RowCount; i++)
            {
                for (j = 0; j < mappinggrid.ColumnCount; j++)
                {
                    if (mappinggrid.Rows[i].Cells[j].Value == null)
                    {
                        tableout.Append("NULL , ");
                    }
                    else
                    {
                        tableout.Append(mappinggrid.Rows[i].Cells[j].Value.ToString() + ", ");
                    }
                }
                tableout.Append("\n");
            }
            tableout.Append("\n rows " + mappinggrid.RowCount);
            tableout.Append("\n columns " + mappinggrid.ColumnCount);
            MessageBox.Show(tableout.ToString());

        }
        private void users_ok_Click(object sender, EventArgs e)
        {

            //group_result1.Text = tools.ADobjectAttributes("CN=user,CN=schema,CN=configuration,DC=fhchs,DC=edu");
            //group_result1.Text = tools.ADobjectAttributes();

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
        private void users_remove_Click(object sender, EventArgs e)
        {

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
        private void group_cancel_Click(object sender, EventArgs e)
        {

        }
        private void group_ok_Click(object sender, EventArgs e)
        {

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
        private void group_see_test_results_Click(object sender, EventArgs e)
        {

            // Loads of test example calls
            //
            // Grab the current domain controller root
            //
            // group_result1.AppendText(getDomain());
            //
            //NEXT
            //
            // Create a list of users in a ou
            //
            //LinkedList<Dictionary<string,string>> ouUsers =  EnumerateUsersInOU("OU=Atis,OU=FHCHS,DC=FHCHS,DC=EDU");
            //LinkedListNode<Dictionary<string, string>> ouUserNode;
            //ouUserNode = ouUsers.First;
            //while (ouUserNode != null)
            //{
            //    group_result1.AppendText(ouUserNode.Value["sAMAccountName"] + "\n");
            //    ouUserNode = ouUserNode.Next;
            //}
            //
            // NEXT
            //
            // Create a SQL insert statement
            // tools.temp_Table(tools.EnumerateUsersInGroup("CN=_AtisRW,OU=Atis,OU=FHCHS,DC=FHCHS,DC=EDU"), "MikesADTest", "soniswebdatabase", "fhcsvdb");
            int i;
            StopWatch timer = new StopWatch();
            timer.Start();
            groupSyncr.ExecuteGroupSync(groupconfig, tools, log, this);
            timer.Stop();
            MessageBox.Show("bulk " +timer.GetElapsedTimeSecs().ToString());

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



        // UI DIALOG  DATA ENTRY EVENTS FOR CONFIGURATION TAB
        private void test_data_source_Click(object sender, EventArgs e)
        {
            SqlConnection sqlConn = new SqlConnection("Data Source=" + DBserver.Text.ToString() + ";Initial Catalog=" + Catalog.Text.ToString() + ";Integrated Security=SSPI;");
            try
            {
                sqlConn.Open();
                groupconfig.DBCatalog = Catalog.Text.ToString();
                userconfig.DBCatalog = Catalog.Text.ToString();
                groupconfig.DataServer = DBserver.Text.ToString();
                userconfig.DataServer = DBserver.Text.ToString();
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
        private void Database_or_AD_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (usermap_Database_or_AD.Text == "Database")
            {
                userMapping_UserData.Visible = true;
                OUdata.Visible = false;
            }
            if (usermap_Database_or_AD.Text == "Active Directory OU")
            {
                userMapping_UserData.Visible = false;
                OUdata.Visible = true;
            }
        }
        private void move_delete_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (usermap_move_delete.Text == "Delete")
            {
                move_users.Visible = false;
            }
            if (usermap_move_delete.Text == "Move")
            {
                move_users.Visible = true;
            }
        }
        private void usermap_active_disabled_SelectedIndexChanged(object sender, EventArgs e)
        {
            usermapping.ActiveDisabled = usermap_active_disabled.Text;
        }
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
                bob.SubItems[1].Text = userMapping_DBServerName.Text;
                bob.SubItems[2].Text = userMapping_DatabaseName.SelectedIndex.ToString();
                bob.SubItems[3].Text = usermapping_user_table_view.SelectedIndex.ToString();
                bob.SubItems[4].Text = usermapping_user_source.SelectedIndex.ToString();
                bob.SubItems[5].Text = usermapping_user_sAMAccountName.SelectedIndex.ToString();
                bob.SubItems[6].Text = usermapping_user_where.Text;
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
                    bob.Text = userMapping_factName.Text;
                    bob.Name = userMapping_factName.Text;
                    bob.SubItems.Add(userMapping_DBServerName.Text);
                    bob.SubItems.Add(userMapping_DatabaseName.SelectedIndex.ToString());
                    bob.SubItems.Add(usermapping_user_table_view.SelectedIndex.ToString());
                    bob.SubItems.Add(usermapping_user_source.SelectedIndex.ToString());
                    bob.SubItems.Add(usermapping_user_sAMAccountName.SelectedIndex.ToString());
                    bob.SubItems.Add(usermapping_user_where.Text);

                    if (bob.Text.Trim().Length == 0 || userMapping_factList.Items.ContainsKey(bob.Text))
                    {
                        MessageBox.Show("I'm Sorry but every fact is required to have a name that is unique.");
                    }
                    else
                    {
                        userMapping_factList.Items.Add(bob);
                        userMapping_fact_Add_Edit.Text = "Add";
                        userMapping_factName.Text = "";
                        userMapping_DBServerName.Text = "";
                        userMapping_DatabaseName.DataSource = null;
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
            userMapping_DBServerName.Text = "";
            userMapping_DatabaseName.DataSource = null;
            usermapping_user_table_view.SelectedIndex = -1;
            usermapping_user_source.DataSource = null;
            usermapping_user_sAMAccountName.DataSource = null;
            usermapping_user_where.Text = "";

            if (userMapping_factList.SelectedItems.Count == 1)
            {
                userMapping_factDelete.Enabled = true;

                userMapping_fact_Add_Edit.Text = "Edit/Save";
                ListViewItem bob = userMapping_factList.SelectedItems[0];
                userMapping_factName.Text                     = bob.Text;
                userMapping_DBServerName.Text                 = bob.SubItems[1].Text;
                userMapping_DatabaseName.SelectedIndex        = Convert.ToInt32(bob.SubItems[2].Text);
                usermapping_user_table_view.SelectedIndex     = Convert.ToInt32(bob.SubItems[3].Text);
                usermapping_user_source.SelectedIndex         = Convert.ToInt32(bob.SubItems[4].Text);
                usermapping_user_sAMAccountName.SelectedIndex = Convert.ToInt32(bob.SubItems[5].Text);
                usermapping_user_where.Text                   = bob.SubItems[6].Text;
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
        }

        // custom AD fields 
    }
}
