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
// UI events call execution methods\
// Test SVn line



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
        //public struct userObject
        //{
        //    public string sAMAccountName;
        //    public string CN;
        //    public Dictionary<string, string> properties;
        //}
        //public struct groupObject
        //{
        //    public string sAMAccountName;
        //    public string CN;
        //    public Dictionary<string, string> properties;
        //}
        public class logFile
        {
            private List<string> logtransactions;
            private List<string> logterrors;
            public logFile()
            {
                logtransactions = new List<string>();
                logterrors = new List<string>();
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

        }
        public class groupSynch
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
            private String configUser_CN;
            private String configUser_table_view;
            private String configUser_sAMAccount;
            private String configUser_dbTable;
            private String configUser_where;
            private String configDataServer;
            private String configDBCatalog;
            private String configprogress;

            // constructor creates blank strings
            public groupSynch()
            {
                configBaseGroupOU = "";
                configBaseUserOU = "";
                configNotes = "";
                configGroup_CN = "";
                configGroup_table_view = "";
                configGroup_sAMAccount = "";
                configGroup_dbTable = "";
                configGroup_where = "";
                configUser_CN = "";
                configUser_table_view = "";
                configUser_sAMAccount = "";
                configUser_dbTable = "";
                configUser_where = "";
                configDataServer = "";
                configDBCatalog = "";
            }
            public groupSynch(Dictionary<string, string> dictionary)
            {
                configBaseGroupOU = dictionary[configBaseGroupOU].ToString();
                configBaseUserOU = dictionary[configBaseUserOU].ToString();
                configNotes = dictionary[configNotes].ToString();
                configGroup_CN = dictionary[configGroup_CN].ToString();
                configGroup_table_view = dictionary[configGroup_table_view].ToString();
                configGroup_sAMAccount = dictionary[configGroup_sAMAccount].ToString();
                configGroup_dbTable = dictionary[configGroup_dbTable].ToString();
                configGroup_where = dictionary[configGroup_where].ToString();
                configUser_CN = dictionary[configUser_CN].ToString();
                configUser_table_view = dictionary[configUser_table_view].ToString();
                configUser_sAMAccount = dictionary[configUser_sAMAccount].ToString();
                configUser_dbTable = dictionary[configUser_dbTable].ToString();
                configUser_where = dictionary[configUser_where].ToString();
                configDataServer = dictionary[configDataServer].ToString();
                configDBCatalog = dictionary[configDBCatalog].ToString();
            }

            public void load(Dictionary<string, string> dictionary)
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
                dictionary.TryGetValue("configUser_CN", out configUser_CN);
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
                returnvalue.Add("configUser_CN", configUser_CN);
                returnvalue.Add("configUser_table_view", configUser_table_view);
                returnvalue.Add("configUser_sAMAccount", configUser_sAMAccount);
                returnvalue.Add("configUser_dbTable", configUser_dbTable);
                returnvalue.Add("configUser_where", configUser_where);
                returnvalue.Add("configDataServer", configDataServer);
                returnvalue.Add("configDBCatalog", configDBCatalog);
                return returnvalue;
            }
            public groupSynch Clone()
            {
                groupSynch retunvalue = new groupSynch();
                retunvalue.load(this.ToDictionary());
                return retunvalue;
            }

        }
        public class userSynch
        {


            private String configBaseUserOU;
            private String configNotes;
            private String configOU_CN;
            private String configOU_table_view;
            private String configOU_sAMAccount;
            private String configOU_dbTable;
            private String configOU_where;
            private String configUser_CN;
            private String configUser_table_view;
            private String configUser_sAMAccount;
            private String configUser_dbTable;
            private String configUser_where;
            private String configDataServer;
            private String configDBCatalog;

            // constructor creates a blank instance
            public userSynch()
            {
                configBaseUserOU = "";
                configNotes = "";
                configOU_CN = "";
                configOU_table_view = "";
                configOU_sAMAccount = "";
                configOU_dbTable = "";
                configOU_where = "";
                configUser_CN = "";
                configUser_table_view = "";
                configUser_sAMAccount = "";
                configUser_dbTable = "";
                configUser_where = "";
                configDataServer = "";
                configDBCatalog = "";
            }
            public void load(Dictionary<string, string> dictionary)
            {
                dictionary.TryGetValue("configBaseUserOU", out configBaseUserOU);
                dictionary.TryGetValue("configNotes", out configNotes);
                dictionary.TryGetValue("configOU_CN", out configOU_CN);
                dictionary.TryGetValue("configOU_table_view", out configOU_table_view);
                dictionary.TryGetValue("configOU_sAMAccount", out configOU_sAMAccount);
                dictionary.TryGetValue("configOU_dbTable", out configOU_dbTable);
                dictionary.TryGetValue("configOU_where", out configOU_where);
                dictionary.TryGetValue("configUser_CN", out configUser_CN);
                dictionary.TryGetValue("configUser_table_view", out configUser_table_view);
                dictionary.TryGetValue("configUser_sAMAccount", out configUser_sAMAccount);
                dictionary.TryGetValue("configUser_dbTable", out configUser_dbTable);
                dictionary.TryGetValue("configUser_where", out configUser_where);
                dictionary.TryGetValue("configDataServer", out configDataServer);
                dictionary.TryGetValue("configDBCatalog", out configDBCatalog);
            }

            // accessors for properties

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
            public String OU_CN
            {
                get
                {
                    return configOU_CN;
                }
                set
                {
                    configOU_CN = value;
                }
            }
            public String OU_table_view
            {
                get
                {
                    return configOU_table_view;
                }
                set
                {
                    configOU_table_view = value;
                }
            }
            public String OU_sAMAccount
            {
                get
                {
                    return configOU_sAMAccount;
                }
                set
                {
                    configOU_sAMAccount = value;
                }
            }
            public String OU_dbTable
            {
                get
                {
                    return configOU_dbTable;
                }
                set
                {
                    configOU_dbTable = value;
                }
            }
            public String OU_where
            {
                get
                {
                    return configOU_where;
                }
                set
                {
                    configOU_where = value;
                }
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

            // output to a dictionary list
            public Dictionary<string, string> ToDictionary()
            {
                Dictionary<string, string> returnvalue = new Dictionary<string, string>();
                returnvalue.Add("configBaseUserOU", configBaseUserOU);
                returnvalue.Add("configNotes", configNotes);
                returnvalue.Add("configOU_CN", configOU_CN);
                returnvalue.Add("configOU_table_view", configOU_table_view);
                returnvalue.Add("configOU_sAMAccount", configOU_sAMAccount);
                returnvalue.Add("configOU_dbTable", configOU_dbTable);
                returnvalue.Add("configOU_where", configOU_where);
                returnvalue.Add("configUser_CN", configUser_CN);
                returnvalue.Add("configUser_table_view", configUser_table_view);
                returnvalue.Add("configUser_sAMAccount", configUser_sAMAccount);
                returnvalue.Add("configUser_dbTable", configUser_dbTable);
                returnvalue.Add("configUser_where", configUser_where);
                returnvalue.Add("configDataServer", configDataServer);
                returnvalue.Add("configDBCatalog", configDBCatalog);
                return returnvalue;
            }
        }
        public class userStateChange
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
            public userStateChange()
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
        public class toolset
        {

            //Functions

            // all functions pulling from AD are limited to 1500 results or immediate failure to fix this things need to be paged max page size 1000
            
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
            public string getDomain()
            {
                using (Domain d = Domain.GetCurrentDomain())
                using (DirectoryEntry entry = d.GetDirectoryEntry())
                {
                    return entry.Path;
                }
            }
            public DataTable EnumerateUsersInOUDataTable(string OuDN)
            {
                // note does not handle special/illegal characters for AD
                // RETURNS ALL USERS IN AN OU NO MATTER HOW DEEP
                DataTable returnvalue = new DataTable();
                DataRow row;

                // bind to the OU you want to enumerate
                DirectoryEntry deOU = new DirectoryEntry("LDAP://" + OuDN);

                // create a directory searcher for that OU
                DirectorySearcher dsUsers = new DirectorySearcher(deOU);

                // set the filter to get just the users
                dsUsers.Filter = "(&(objectClass=user)(objectCategory=Person))";

                // add the attributes you want to grab from the search
                // COULD OVERLOAD METHOD AND CHANGE OUT FOR A FOREACH AND GRAB PROPS FROM AN ARRAY
                dsUsers.PropertiesToLoad.Add("sAMAccountName");

                // grab the users and do whatever you need to do with them 
                row = returnvalue.NewRow();
                foreach (SearchResult oResult in dsUsers.FindAll())
                {
                    //generate the array list with the user sam accounts
                    // COULD CHANGE OUT FOR A FOR EACH AND GRAB PROPS FROM AN ARRAY
                    row[0] = oResult.Properties["sAMAccountName"][0].ToString();
                    returnvalue.Rows.Add(row);
                    row = returnvalue.NewRow();
                }
                return returnvalue;
            }
            public DataTable EnumerateUsersInGroupDataTable(string groupDN)
            {
                // note does not handle special/illegal characters for AD
                // groupDN "CN=Sales,OU=test,DC=Fabrikam,DC=COM"
                // returns FQDN "CN=user,OU=test,DC=Fabrikam,DC=COM" & group "CN=Sales,OU=test,DC=Fabrikam,DC=COM" of users in group 
                DataTable returnvalue = new DataTable();
                DataRow row;
                DirectoryEntry group = new DirectoryEntry("LDAP://" + System.Web.HttpUtility.UrlEncode(groupDN).Replace("+", " ").Replace("*", "%2A"));
                DirectorySearcher groupUsers = new DirectorySearcher(group);
                returnvalue.Columns.Add("FQDN");
                returnvalue.Columns.Add("group");
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
            public DataTable EnumerateGroupsInOUDataTable(string OuDN, ArrayList returnProperties, string table)
            {
                // note does not handle special/illegal characters for AD
                int i;
                int count = returnProperties.Count;
                DataTable returnvalue = new DataTable();
                DataRow row;
                // bind to the OU you want to enumerate
                DirectoryEntry deOU = new DirectoryEntry("LDAP://" + OuDN);                

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
            public bool createOURecursive(string ou)
            {
                if (Exists(ou) == true)
                {
                    return true;
                }
                else
                {
                    createOURecursive(ou.Substring(ou.IndexOf(",") + 1));
                    CreateOU(ou.Substring(ou.IndexOf(",") + 1), ou.Remove(ou.IndexOf(",")).Substring(ou.IndexOf("=") + 1));
                    return true;
                }
            }
            public void CreateGroup(string ouPath, Dictionary<string, string> properties)
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
                    }
                    else
                    { 
                        MessageBox.Show("CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + " group already exists from adding");
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message.ToString() + "issue create group LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath);
                }
            }
            public void UpdateGroup(string ouPath, Dictionary<string, string> properties)
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
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message.ToString() + "issue updating group LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath);
                    }
                }
                else
                { MessageBox.Show(ouPath + " group already exists from updating");
                }
            }
            public void DeleteGroup(string ouPath, string name)
            {
                if (DirectoryEntry.Exists("LDAP://CN=" + System.Web.HttpUtility.UrlEncode(name).Replace("+", " ").Replace("*", "%2A") + "," + ouPath))
                {
                    try
                    {
                        DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                        DirectoryEntry group = new DirectoryEntry("LDAP://CN=" + System.Web.HttpUtility.UrlEncode(name).Replace("+", " ").Replace("*", "%2A") + "," + ouPath);
                        entry.Children.Remove(group);
                        group.CommitChanges();
                    }
                    catch (Exception e)
                    {
                        // MessageBox.Show(e.Message.ToString() + " error deleting LDAP://CN=" + name + "," + ouPath );
                    }
                }
                else
                {
                    // MessageBox.Show("LDAP://CN=" + name + "," + ouPath);
                }
            }
            public void CreateOU(string ouPath, string name)
            {
                //needs parent OU present to work
                if (!DirectoryEntry.Exists("LDAP://OU=" + name + "," + ouPath))
                {
                    try
                    {
                        DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                        DirectoryEntry OU = entry.Children.Add("OU=" + name, "organizationalUnit");
                        //                   OU.Properties["sAmAccountName"].Value = name;
                        OU.CommitChanges();
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message.ToString() + "create ou LDAP://OU=" + name + "," + ouPath);
                    }
                }
                else
                { // MessageBox.Show("LDAP://OU=" + name + "," + ouPath + " already exists"); 
                }
            }
            public void DeleteOU(string ouPath, string name)
            {
                //needs parent OU present to work
                if (!DirectoryEntry.Exists("LDAP://OU=" + name + "," + ouPath))
                {
                    try
                    {
                        DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                        entry.DeleteTree();
                    }
                    catch (Exception e)
                    {
                        // MessageBox.Show(e.Message.ToString() + "create ou LDAP://OU=" + name + "," + ouPath);
                    }
                }
                else
                { // MessageBox.Show("LDAP://OU=" + name + "," + ouPath + " already exists"); 
                }
            }
            public void AddUserToGroup(string userDn, string groupDn)
            {
                try
                {
                    if (Exists(userDn))
                    {
                        DirectoryEntry dirEntry = new DirectoryEntry("LDAP://" + groupDn);
                        dirEntry.Properties["member"].Add(userDn);
                        dirEntry.CommitChanges();
                        dirEntry.Close();
                    }
                }
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    // MessageBox.Show(E.Message.ToString() + " error adding " + userDn + " to LDAP://" + groupDn);

                }
            }
            public void RemoveUserFromGroup(string userDn, string groupDn)
            {
                try
                {
                    DirectoryEntry dirEntry = new DirectoryEntry("LDAP://" + groupDn);
                    try
                    {
                        dirEntry.Properties["member"].Remove(userDn);
                    }
                    catch (System.DirectoryServices.DirectoryServicesCOMException E)
                    { }
                    dirEntry.CommitChanges();
                    dirEntry.Close();
                }
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    // MessageBox.Show(E.Message.ToString() + " error removing " + userDn + " from LDAP://" + groupDn);

                }
            }
            // move user OU
            public bool createUserAccount(string parentOUDN, string samName, string userPassword, string firstName, string lastName)
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
            public bool disableUser(string sAMAccountName, string LdapDomain)
            {
                string userDN;
                userDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, LdapDomain);
                DirectoryEntry usr = new DirectoryEntry(userDN);
                int val = (int)usr.Properties["userAccountControl"].Value;
                usr.Properties["userAccountControl"].Value = val | (int)accountFlags.ADS_UF_ACCOUNTDISABLE;
                usr.CommitChanges();
                return false;
            }
            public bool enableUser(string sAMAccountName, string LdapDomain)
            {
                string userDN;
                userDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, LdapDomain);
                DirectoryEntry usr = new DirectoryEntry(userDN);
                int val = (int)usr.Properties["userAccountControl"].Value;
                usr.Properties["userAccountControl"].Value = val | (int)accountFlags.ADS_UF_ACCOUNTDISABLE;
                usr.CommitChanges();
                return false;
            }
            public string create_Table(DataTable data, string table, SqlConnection sqlConn)
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
            public string append_to_Table(DataTable data, string table, SqlConnection sqlConn)
            {
                // copy data into table
                SqlBulkCopy sbc = new SqlBulkCopy(sqlConn);
                sbc.DestinationTableName = table;
                sbc.WriteToServer(data);
                sbc.Close();
                return table;
            }
            public bool deleteUserAccount(string sAMAccountName, string LdapDomain)
            {
                string user;

                user = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, LdapDomain);
                DirectoryEntry ent = new DirectoryEntry(user);
                ent.DeleteTree();
                return false;
            }
            

            public SqlDataReader queryNotExists(string Table1, string Table2, SqlConnection sqlConn, string pkey1, string pkey2)
            {
                // finds items in table1 who do not exist in table2 and returns them
                // SqlCommand sqlComm = new SqlCommand("Select Table1.* Into #Table3ADTransfer From " + Table1 + " AS Table1, " + Table2 + " AS Table2 Where Table1." + pkey1 + " = Table2." + pkey2 + " And Table2." + pkey2 + " is null", sqlConn);
                SqlCommand sqlComm = new SqlCommand("SELECT uptoDate.* FROM " + Table1 + " uptoDate LEFT OUTER JOIN " + Table2 + " outofDate ON outofDate." + pkey2 + " = uptoDate." + pkey1 + " WHERE outofDate." + pkey2 + " IS NULL;", sqlConn);
                // create the command object
                SqlDataReader r = sqlComm.ExecuteReader();
                return r;
            }
            public SqlDataReader checkUpdate(string table1, string table2, string pkey1, string pkey2, ArrayList compareFields1, ArrayList compareFields2, SqlConnection sqlConn)
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

            // additional stuff
            public bool setUserExpiration(int days, string LdapDomain, string sAMAccountName)
            {
                string usrDN;
                usrDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, LdapDomain);
                DirectoryEntry usr = new DirectoryEntry(usrDN);
                Type type = usr.NativeObject.GetType();
                Object adsNative = usr.NativeObject;
                string formattedDate;

                // Calculating the new date
                DateTime yesterday = DateTime.Today.AddDays(days);
                formattedDate = yesterday.ToString("dd/MM/yyyy");

                type.InvokeMember("AccountExpirationDate", BindingFlags.SetProperty, null, adsNative, new object[] { formattedDate });
                usr.CommitChanges();
                return true;
            }
            public ArrayList setMultiPropertyUser(LinkedList<Dictionary<string, string>> userList, ArrayList propertyArray, string LdapDomain)
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
                    usrDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, LdapDomain);
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
            public static string ReplaceEscapeChars(string str)
            {
                //If the string is null
                if (str == null)
                    return str;

                //If the string is empty
                if (str == "")
                    return str;

                //Replaces single quote (') with two (2) single quotes ('')
                //solves the problem of inserting, updating or selecting a text with single quote (')
                //i.e.: Cox's Bazar, World's economy etc.
                str = str.Replace("'", "''");
                return str;
            }
        }

        public class objectADSqlsyncGroup
        {
            public void executeBulkcopy(groupSynch groupsyn, toolset tools, logFile log, Form1 gui)
            {
                //string debug = "";
                //SqlDataReader debugreader;
                //ArrayList debuglist = new ArrayList();
                //int debugfieldcount;
                //string debugrecourdcount;
                //int i;
                //StopWatch time = new StopWatch();
                //SqlCommand sqldebugComm;

                string groupapp = groupsyn.Group_Append;
                string groupOU = groupsyn.BaseGroupOU;
                string sAMAccountName = "";
                string description = "";
                string sqlgroupsTable = "#FHC_SQLgroupsTable";
                string ADgroupsTable = "#FHC_ADgroupsTable";
                string DC = groupOU.Substring(groupOU.IndexOf("DC"));
                string groupDN;
                string groupsTable;
                SqlDataReader add;
                SqlDataReader delete;
                SqlDataReader update;
                ArrayList ADupdateKeys = new ArrayList();
                ArrayList SQLupdateKeys = new ArrayList();
                ArrayList fields = new ArrayList();
                DataTable groupsDataTable = new DataTable();
                Dictionary<string, string> groupObject = new Dictionary<string, string>();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + groupsyn.DataServer + ";Initial Catalog=" + groupsyn.DBCatalog + ";Integrated Security=SSPI;");



                sqlConn.Open();
                // Setup the OU for the program
                tools.createOURecursive("OU=" + groupapp + "," + groupOU);

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
                ADupdateKeys.Add("description");
                ADupdateKeys.Add("CN");


                // grab groups from AD
                //time.Start();
                groupsDataTable = tools.EnumerateGroupsInOUDataTable("OU=" + groupapp + "," + groupOU, ADupdateKeys, ADgroupsTable);
                //time.Stop();
                gui.Refresh();
                //MessageBox.Show("got " + groupsLinkedList.Count + "groups from ou in " + time.GetElapsedTime());
                // insert groups from AD into a temp table
                if (groupsDataTable.Rows.Count > 0)
                {
                    //time.Start();
                    groupsTable = tools.create_Table(groupsDataTable, ADgroupsTable, sqlConn);
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
                    add = tools.queryNotExists(sqlgroupsTable, groupsTable, sqlConn, groupsyn.Group_CN, ADupdateKeys[1].ToString());

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
                        tools.CreateGroup("OU=" + groupapp + "," + groupOU, groupObject);
                        log.transactions.Add("Group added ;" + sAMAccountName + ",OU=" + groupapp + "," + groupOU + ";" + description);
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
                    delete = tools.queryNotExists(groupsTable, sqlgroupsTable, sqlConn, ADupdateKeys[1].ToString(), groupsyn.Group_CN);
                    // delete groups in AD
                   // i = 0;
                    while (delete.Read())
                    {
                       // i++;
                        tools.DeleteGroup("OU=" + groupapp + "," + groupOU, (string)delete[ADupdateKeys[1].ToString()].ToString().Trim());
                        log.transactions.Add("Group deleted ;" + (string)delete[ADupdateKeys[1].ToString()].ToString().Trim() + ",OU=" + groupapp + groupOU);
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
                    SQLupdateKeys.Add(groupsyn.Group_sAMAccount);
                    SQLupdateKeys.Add(groupsyn.Group_CN);
                    //time.Start();
                    // update assumes the both ADupdateKeys and SQLupdateKeys have the same fields, listed in the same order check  call to EnumerateGroupsInOU if this is wrong should be sAMAccountName, CN matching the SQL order
                    update = tools.checkUpdate(sqlgroupsTable, groupsTable, groupsyn.Group_CN, ADupdateKeys[1].ToString(), SQLupdateKeys, ADupdateKeys, sqlConn);
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
                            tools.UpdateGroup("OU=" + groupapp + "," + groupOU, groupObject);
                            log.transactions.Add("Group update ; " + sAMAccountName + ",OU=" + groupapp + "," + groupOU + ";" + description);
                        }
                        else
                        {
                            // find it its on the server somewhere we will log the exception
                            groupDN = tools.GetObjectDistinguishedName(objectClass.group, returnType.distinguishedName, groupObject["CN"], DC);
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
                        tools.CreateGroup("OU=" + groupapp + "," + groupOU, groupObject);
                        log.transactions.Add("Group added ;" + groupObject["sAMAccountName"] + ",OU=" + groupapp + "," + groupOU + ";" + groupObject["description"]);

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
                    sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + groupsyn.User_sAMAccount + ") AS " + groupsyn.User_sAMAccount + ", RTRIM(" + groupsyn.User_CN + ") AS " + groupsyn.User_CN + " INTO " + sqlgroupMembersTable + " FROM " + groupsyn.User_dbTable, sqlConn);
                }
                else
                {
                    sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + groupsyn.User_sAMAccount + ") AS " + groupsyn.User_sAMAccount + ", RTRIM(" + groupsyn.User_CN + ") AS " + groupsyn.User_CN + " INTO " + sqlgroupMembersTable + " FROM " + groupsyn.User_dbTable + " WHERE " + groupsyn.User_where, sqlConn);
                }
                sqlComm.ExecuteNonQuery();

                // populate datatable with users from AD groups by looping thru the list of groups from SQL and loading the cross referenced AD group members
                sqlComm = new SqlCommand("select " + groupsyn.Group_CN + " FROM " + sqlgroupsTable, sqlConn);
                sqlgroups = sqlComm.ExecuteReader();
                while (sqlgroups.Read())
                {
                       // hopefully merge acts as an append
                    ADusers.Merge(tools.EnumerateUsersInGroupDataTable("CN=" + (string)sqlgroups[0] + ",OU=" + groupapp + "," + groupOU));
                }
                sqlgroups.Close();

                // make the temp table for ou comparisons
                tools.create_Table(ADusers, ADgroupMembersTable, sqlConn);
                sqlComm = new SqlCommand("select count(" + ADusers.Columns[0].ColumnName + ") FROM " + ADgroupMembersTable, sqlConn);
                MessageBox.Show("table " + ADgroupMembersTable + " has " + sqlComm.ExecuteScalar().ToString() + " records \n ");
                
                // compare and add/remove
                add = tools.queryNotExists(sqlgroupMembersTable, ADgroupMembersTable, sqlConn, groupsyn.User_CN, ADusers.Columns[0].ColumnName);
                while (add.Read())
                {
                    tools.AddUserToGroup((string)add[0], (string)add[1]);
                    log.transactions.Add("User added ;" + (string)add[0] + ",OU=" + groupapp + "," + groupOU + ";" + description);
                    groupObject.Clear();
                }
                add.Close();

                delete = tools.queryNotExists(sqlgroupMembersTable, ADgroupMembersTable, sqlConn, groupsyn.Group_CN, ADupdateKeys[1].ToString());
                // delete groups in AD
                while (delete.Read())
                {

                    tools.RemoveUserFromGroup((string)delete[0], (string)delete[1]);
                    log.transactions.Add("User removed ;" + (string)delete[ADupdateKeys[1].ToString()].ToString().Trim() + ",OU=" + groupapp + groupOU);

                }
                delete.Close();
                sqlConn.Close();
            }
        }

        // create objects to hold save data
        groupSynch groupconfig = new groupSynch();
        userSynch userconfig = new userSynch();
        executionOrder execution = new executionOrder();
        userStateChange usermapping = new userStateChange();
        toolset tools = new toolset();
        logFile log = new logFile();
        objectADSqlsyncGroup groupSyncr = new objectADSqlsyncGroup();

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
        //        sqlComm = new SqlCommand("SELECT " + groupconfig.User_sAMAccount + ", " + groupconfig.User_CN + " FROM " + groupconfig.User_dbTable, sqlConn);
        //    }
        //    else
        //    {
        //        sqlComm = new SqlCommand("SELECT " + groupconfig.User_sAMAccount + ", " + groupconfig.User_CN + " FROM " + groupconfig.User_dbTable + " WHERE " + groupconfig.User_where, sqlConn);
        //    }
        //    r = sqlComm.ExecuteReader();

        //    // interate thru a recordset based on query generated from text and generate the linked list of dictionary for diff
        //    while (r.Read())
        //    {
        //        UsergroupDictionary = new Dictionary<string, string>();
        //        //  groupconfig.User_CN holds the value for the cross refernce aginst the group CN 
        //        UsergroupDictionary.Add("sAMAccountName", (string)r[groupconfig.User_sAMAccount].ToString().Trim());
        //        UsergroupDictionary.Add("CN", (string)r[groupconfig.User_CN].ToString().Trim() + groupconfig.Group_Append);
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
                users_user_CN.DataSource = columnList;
                users_user_sAMAccountName.DataSource = columnList.Clone();
            }
            else
            {
                MessageBox.Show("Please select table or view");
            }
        }
        private void usersMap_user_CN_SelectedIndexChanged(object sender, EventArgs e)
        {
            userconfig.User_CN = users_user_CN.Text.ToString();
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
        private void users_ou_Table_View_SelectedIndexChanged(object sender, EventArgs e)
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
                if (users_ou_Table_View.Text.ToLower() == "table")
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

                userconfig.OU_table_view = users_ou_Table_View.Text.ToString();
                users_ou_source.DataSource = tableList;
            }
            else
            {
                MessageBox.Show("Please set the dataserver and catalog");
            }
        }
        private void users_ou_source_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (userconfig.DBCatalog != "" && userconfig.DataServer != "")
            {
                //populates columns dialog with columns depending on the results of a query
                ArrayList columnList = new ArrayList();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + userconfig.DataServer.ToString() + ";Initial Catalog=" + userconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");

                sqlConn.Open();
                // create the command object
                SqlCommand sqlComm = new SqlCommand("SELECT column_name FROM information_schema.columns WHERE table_name = '" + users_ou_source.Text.ToString() + "'", sqlConn);
                SqlDataReader r = sqlComm.ExecuteReader();
                while (r.Read())
                {
                    columnList.Add((string)r[0].ToString().Trim());
                }
                r.Close();
                sqlConn.Close();

                userconfig.OU_dbTable = users_ou_source.Text.ToString();
                users_ou_CN.DataSource = columnList;
                users_ou_sAMAccountName.DataSource = columnList.Clone();
            }
            else
            {
                MessageBox.Show("Please set the dataserver and catalog");
            }

        }
        private void users_ou_CN_SelectedIndexChanged(object sender, EventArgs e)
        {
            userconfig.OU_CN = users_ou_CN.Text.ToString();
        }
        private void users_ou_sAMAccount_SelectedIndexChanged(object sender, EventArgs e)
        {
            userconfig.OU_sAMAccount = users_ou_sAMAccountName.Text.ToString();
        }
        private void users_ou_where_TextChanged(object sender, EventArgs e)
        {
            userconfig.OU_where = users_ou_where.Text.ToString();
        }

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
                users_user_CN.DataSource = columnList;
                users_user_sAMAccountName.DataSource = columnList.Clone();
            }
            else
            {
                MessageBox.Show("Please select table or view");
            }
        }
        private void users_user_CN_SelectedIndexChanged(object sender, EventArgs e)
        {
            userconfig.User_CN = users_user_CN.Text.ToString();
        }
        private void users_user_sAMAccountName_SelectedIndexChanged(object sender, EventArgs e)
        {
            userconfig.User_sAMAccount = users_user_sAMAccountName.Text.ToString();
        }
        private void users_user_where_TextChanged(object sender, EventArgs e)
        {
            userconfig.User_where = users_user_where.Text.ToString();
        }
        private void users_baseUserOU_TextChanged(object sender, EventArgs e)
        {
            userconfig.BaseUserOU = users_baseUserOU.Text.ToString();
        }
        private void users_mapping_description_TextChanged(object sender, EventArgs e)
        {
            userconfig.Notes = users_mapping_description.Text.ToString();
        }
        // BUTTONS FOR THE TAB
        private void users_see_query_Click(object sender, EventArgs e)
        {
            users_result1.Clear();
            users_result1.AppendText("This is your OU query \n");
            users_result1.AppendText("Select ");
            users_result1.AppendText(userconfig.OU_CN);
            users_result1.AppendText(", ");
            users_result1.AppendText(userconfig.OU_sAMAccount);
            users_result1.AppendText(" From ");
            users_result1.AppendText(userconfig.OU_dbTable);
            if (userconfig.OU_where != string.Empty)
            {
                users_result1.AppendText(" Where ");
                users_result1.AppendText(userconfig.OU_where);
            }
            users_result1.AppendText("\n");

            users_result2.Clear();
            users_result2.AppendText("This is your users query \n");
            users_result2.AppendText("Select ");
            users_result2.AppendText(userconfig.User_CN);
            users_result2.AppendText(", ");
            users_result2.AppendText(userconfig.User_sAMAccount);
            users_result2.AppendText(" From ");
            users_result2.AppendText(userconfig.OU_dbTable);
            if (userconfig.User_where != string.Empty)
            {
                users_result2.AppendText(" Where ");
                users_result2.AppendText(userconfig.User_where);
            }
            users_result2.AppendText("\n");
        }
        private void users_see_test_results_Click(object sender, EventArgs e)
        {

        }
        private void users_Save_button(object sender, EventArgs e)
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

        }
        private void users_ok_Click(object sender, EventArgs e)
        {

        }
        private void users_open_Click(object sender, EventArgs e)
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
            userconfig.load(properties);
            DBserver.Text = userconfig.DataServer;
            userconfig.load(properties);
            Catalog.Text = userconfig.DBCatalog;
            userconfig.load(properties);
            users_ou_Table_View.Text = userconfig.OU_table_view;
            userconfig.load(properties);
            users_ou_source.Text = userconfig.OU_dbTable;
            userconfig.load(properties);
            users_ou_CN.Text = userconfig.OU_CN;
            userconfig.load(properties);
            users_ou_sAMAccountName.Text = userconfig.OU_sAMAccount;
            userconfig.load(properties);
            users_ou_where.Text = userconfig.User_table_view;

            users_user_Table_View.Text = userconfig.User_table_view;
            userconfig.load(properties);
            users_user_source.Text = userconfig.User_dbTable;
            userconfig.load(properties);
            users_user_CN.Text = userconfig.User_CN;
            userconfig.load(properties);
            users_user_sAMAccountName.Text = userconfig.User_sAMAccount;
            userconfig.load(properties);
            users_user_where.Text = userconfig.User_where;

            users_mapping_description.Text = userconfig.Notes;
            users_baseUserOU.Text = userconfig.BaseUserOU;

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
                group_user_CN.DataSource = columnList;
                group_user_sAMAccountName.DataSource = columnList.Clone();
            }
            else
            {
                MessageBox.Show("Please select table or view");
            }
        }
        private void group_user_CN_SelectedIndexChanged(object sender, EventArgs e)
        {
            groupconfig.User_CN = group_user_CN.Text.ToString();
        }
        private void group_user_sAMAccountName_SelectedIndexChanged(object sender, EventArgs e)
        {
            groupconfig.User_sAMAccount = group_user_sAMAccountName.Text.ToString();
        }
        private void group_user_where_TextChanged(object sender, EventArgs e)
        {
            groupconfig.User_where = group_user_where.Text.ToString();
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
            group_result2.AppendText(groupconfig.User_CN);
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
            groupconfig.load(properties);
            DBserver.Text = groupconfig.DataServer;
            Catalog.Text = groupconfig.DBCatalog;
            group_group_Table_View.Text = groupconfig.Group_table_view;
            groupconfig.load(properties);
            group_group_source.Text = groupconfig.Group_dbTable;
            groupconfig.load(properties);
            group_group_CN.Text = groupconfig.Group_CN;
            groupconfig.load(properties);
            group_group_sAMAccountName.Text = groupconfig.Group_sAMAccount;
            groupconfig.load(properties);
            group_group_where.Text = groupconfig.Group_where;
            groupconfig.load(properties);
            group_group_prepend.Text = groupconfig.Group_Append;

            group_user_Table_View.Text = groupconfig.User_table_view;
            groupconfig.load(properties);
            group_user_source.Text = groupconfig.User_dbTable;
            groupconfig.load(properties);
            group_user_CN.Text = groupconfig.User_CN;
            groupconfig.load(properties);
            group_user_sAMAccountName.Text = groupconfig.User_sAMAccount;
            groupconfig.load(properties);
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
            groupSyncr.executeBulkcopy(groupconfig, tools, log, this);
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


            for (i = 0; i < log.transactions.Count; i++)
            {
                group_result1.AppendText(log.transactions[i].ToString() + "\n");
            }
            for (i = 0; i < log.errors.Count; i++)
            {
                group_result2.AppendText(log.errors[i].ToString() + "\n");
            }
            // groupSyncr.execute(groupconfig, tools, log);
            // users_result1.Text log.transactions.ToString();
            // users_result2.Text = log.errors.ToString();
            // MessageBox.Show("compelete");
        }



        // UI DIALOG  DATA ENTRY EVENTS FOR CONFIGURATION TAB
        private void test_data_source_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection("Data Source=" + DBserver.Text.ToString() + ";Initial Catalog=" + Catalog.Text.ToString() + ";Integrated Security=SSPI;");
            try
            {
                con.Open();
                groupconfig.DBCatalog = Catalog.Text.ToString();
                userconfig.DBCatalog = Catalog.Text.ToString();
                groupconfig.DataServer = DBserver.Text.ToString();
                userconfig.DataServer = DBserver.Text.ToString();
                con.Close();
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
                UserData.Visible = true;
                OUdata.Visible = false;
            }
            if (usermap_Database_or_AD.Text == "Active Directory OU")
            {
                UserData.Visible = false;
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
            usermapping.User_dbTable = usermap_user_source.Text;
        }
        private void usermap_user_sAMAccountName_SelectedIndexChanged(object sender, EventArgs e)
        {
            usermapping.User_sAMAccount = usermap_user_sAMAccountName.Text;
        }
        private void usermap_user_table_view_SelectedIndexChanged(object sender, EventArgs e)
        {
            usermapping.User_table_view = usermap_user_table_view.Text;
        }
        private void usermap_user_where_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {
            usermapping.User_where = usermap_user_where.Text;
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
    }
}