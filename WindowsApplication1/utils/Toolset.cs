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
using System.Data.SqlTypes;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Windows.Forms;

//Get these from http://code.google.com/p/google-gdata/
using Google.GData.Apps;
using Google.GData.Apps.GoogleMailSettings;
using Google.GData.Client;

using WindowsApplication1;
using WindowsApplication1.utils;


// outstanding issues
// send as use fix from google forums
// update gmail failing to use middle name properly
// ensure nicknames are genereated properly
// allow for nulls in blank fields to be matching
// unique table naming for multiple instances running at once


// Wish list
// preview area




namespace WindowsApplication1.utils
{
    public enum objectClass
    {
        user, group, computer, organizationalunit
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
    public class LogFile
    {
        private List<string> logtransactions;
        private List<string> logterrors;
        private List<string> logwarnings;
        private List<string> logqueries;
        private DataRow row;
        private DateTime nowstamp = new DateTime();
        public LogFile()
        {
            logtransactions = new List<string>();
            logterrors = new List<string>();
            logwarnings = new List<string>();
            logqueries = new List<string>();
        }

        private DataTable logtrn = new DataTable();
        public void initiateTrn()
        {
            try
            {
                logtrn.Columns.Add("Message");
                logtrn.Columns.Add("Type");
                logtrn.Columns.Add("Timestamp", System.Type.GetType("System.DateTime"));
            }
            catch (Exception)
            {
               
            }
        }

        public DataTable logTrns
        {
            get
            {
                return logtrn;
            }
            set
            {
                logtrn = value;
            }
        }
        public void addTrn(string message, string type)
        {
            nowstamp = DateTime.Now;
            row = logtrn.NewRow();
            row[0] = message.ToString();
            row[1] = type.ToString();
            row[2] = nowstamp;
            logtrn.Rows.Add(row);
        }

        public DataTable toDataTableTrn()
        {
            DataTable returnvalue = new DataTable();
            DataRow row;
            int i = 0;
            row = returnvalue.NewRow();
            for (i = 0; i < transactions.Count; i++)
            {
                row[i] = transactions[i].ToString();
            }
            returnvalue.Rows.Add(row);
            row = returnvalue.NewRow();
            return returnvalue;
        }
        public DataTable toDataTableErr()
        {
            DataTable returnvalue = new DataTable();
            DataRow row;
            int i = 0;
            row = returnvalue.NewRow();
            for (i = 0; i < errors.Count; i++)
            {
                row[i] = errors[i].ToString();
            }
            returnvalue.Rows.Add(row);
            row = returnvalue.NewRow();
            return returnvalue;
        }
        public DataTable toDataTableWar()
        {
            DataTable returnvalue = new DataTable();
            DataRow row;
            int i = 0;
            row = returnvalue.NewRow();
            for (i = 0; i < warnings.Count; i++)
            {
                row[i] = warnings[i].ToString();
            }
            returnvalue.Rows.Add(row);
            row = returnvalue.NewRow();
            return returnvalue;
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
        public List<string> queries
        {
            get
            {
                return logqueries;
            }
            set
            {
                logqueries = value;
            }
        }


    }

    // Object datastructure objects for holding data for the execution phase
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
        //private String configGroup_Security_Distrib;
        private String configUser_Group_Reference;
        private String configUser_table_view;
        private String configUser_sAMAccount;
        private String configUser_dbTable;
        private String configUser_where;
        private String configDataServer;
        private String configDBCatalog;
        //private String configprogress;

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
            configGroup_Prepend = "";
            //configGroup_Security_Distrib = "";
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
            configGroup_Prepend = dictionary[configGroup_Prepend].ToString();
            configGroup_where = dictionary[configGroup_where].ToString();
            //configGroup_Security_Distrib = dictionary[configGroup_Security_Distrib].ToString();
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
            //dictionary.TryGetValue("configGroup_Security_Distrib", out configGroup_Security_Distrib);
            dictionary.TryGetValue("configUser_Group_Reference", out configUser_Group_Reference);
            dictionary.TryGetValue("configUser_table_view", out configUser_table_view);
            dictionary.TryGetValue("configUser_sAMAccount", out configUser_sAMAccount);
            dictionary.TryGetValue("configUser_dbTable", out configUser_dbTable);
            dictionary.TryGetValue("configUser_where", out configUser_where);
            dictionary.TryGetValue("configDataServer", out configDataServer);
            dictionary.TryGetValue("configDBCatalog", out configDBCatalog);
        }

        // accessor to properties
        //public String progress
        //{
        //    get
        //    {
        //        return configprogress;
        //    }
        //    set
        //    {
        //        configprogress = value;
        //    }
        //}
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
        //public String Group_Security_Distrib
        //{
        //    get
        //    {
        //        return configGroup_Security_Distrib;
        //    }
        //    set
        //    {
        //        configGroup_Security_Distrib = value;
        //    }
        //}
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
            // returnvalue.Add("configGroup_Security_Distrib", configGroup_Security_Distrib);
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
        private String configSearchScope;
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
        private String configUser_CN;
        //private String configUser_mail;
        private String configUser_table_view;
        private String configUser_dbTable;
        private String configUser_where;
        private String configDataServer;
        private String configDBCatalog;
        private String configUserHoldingTank;
        private String configUser_password;
        private Boolean configUser_UpdateOnly;
        //private String configUserEmailDomain;
        private DataTable configCustoms = new DataTable();
        private string custom = "";
        private string boolconv = "";
        int i = 0;
        private DataRow row;



        // constructor creates a blank instance
        public UserSynch()
        {
            configSearchScope = "";
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
            configUser_CN = "";
            configUser_dbTable = "";
            //configUser_mail = "";
            configUser_where = "";
            configDataServer = "";
            configDBCatalog = "";
            configUserHoldingTank = "";
            configUser_password = "";
            configUser_UpdateOnly = false;
            //configUserEmailDomain = "";
            configCustoms.Columns.Add("ad");
            configCustoms.Columns.Add("sql");
            configCustoms.Columns.Add("static");
        }
        public void Load(Dictionary<string, string> dictionary)
        {
            dictionary.TryGetValue("configSearchScope", out configSearchScope);
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
            dictionary.TryGetValue("configUser_CN", out configUser_CN);
            //dictionary.TryGetValue("configUser_mail", out configUser_mail);
            dictionary.TryGetValue("configUser_dbTable", out configUser_dbTable);
            dictionary.TryGetValue("configUser_where", out configUser_where);
            dictionary.TryGetValue("configDataServer", out configDataServer);
            dictionary.TryGetValue("configDBCatalog", out configDBCatalog);
            dictionary.TryGetValue("configUserHoldingTank", out configUserHoldingTank);
            dictionary.TryGetValue("configUser_password", out configUser_password);
            dictionary.TryGetValue("configUser_UpdateOnly", out boolconv);
            if (boolconv == "True")
            {
                configUser_UpdateOnly = true;
            }
            else
            {
                configUser_UpdateOnly = false;
            }
            dictionary.TryGetValue("configCustoms", out custom);
            //dictionary.TryGetValue("configUserEmailDomain", out configUserEmailDomain);

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
        public String SearchScope
        {
            get
            {
                return configSearchScope;
            }
            set
            {
                configSearchScope = value;
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
        //public String User_mail
        //{
        //    get
        //    {
        //        return configUser_mail;
        //    }
        //    set
        //    {
        //        configUser_mail = value;
        //    }
        //}
        //public String UserEmailDomain
        //{
        //    get
        //    {
        //        return configUserEmailDomain;
        //    }
        //    set
        //    {
        //        configUserEmailDomain = value;
        //    }
        //}
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
        public Boolean UpdateOnly
        {
            get
            {
                return configUser_UpdateOnly;
            }
            set
            {
                configUser_UpdateOnly = value;
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
            returnvalue.Add("configSearchScope", configSearchScope);
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
            returnvalue.Add("configUser_CN", configUser_CN);
            //returnvalue.Add("configUser_mail", configUser_mail);
            returnvalue.Add("configUser_dbTable", configUser_dbTable);
            returnvalue.Add("configUser_where", configUser_where);
            returnvalue.Add("configDataServer", configDataServer);
            returnvalue.Add("configDBCatalog", configDBCatalog);
            returnvalue.Add("configUniversalGroup", configUniversalGroup);
            returnvalue.Add("configUserHoldingTank", configUserHoldingTank);
            returnvalue.Add("configUser_password", configUser_password);
            returnvalue.Add("configUser_UpdateOnly", configUser_UpdateOnly.ToString());
            //returnvalue.Add("configUserEmailDomain", configUserEmailDomain);
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
    public class GmailUsers
    {
        private String configDataServer;
        private String configDBCatalog;

        private String configUser_adminUser;
        private String configUser_adminDomain;
        private String configUser_adminPassword;

        private String configUser_user_Datasource;
        private String configUser_table_view;
        private String configUser_dbTable;
        private String configUser_where;
        private String configUser_user_ad_OU;

        private String configUser_Lname;
        private String configUser_Fname;
        private String configUser_MiddleName;
        private String configUser_StuID;
        private Boolean configUser_password_generate_checkbox;
        private Boolean configUser_password_short_fix_checkbox;
        private String configUser_password;
        private Boolean configUser_levenshtein;

        private Boolean configUser_writeback_AD_checkbox;
        private Boolean configUser_writeback_DB_checkbox;
        private String configUser_writeback_table;
        private String configUser_writeback_primary_key;
        private Boolean configUser_writeback_transfer_email_checkbox;
        private String configUser_writeback_where_clause;
        private String configUser_writeback_email_field;
        private String configUser_writeback_secondary_email_field;
        private String configUser_writeback_ad_OU;
        private string boolconv = "";



        public GmailUsers()
        {
            configDataServer = "";
            configDBCatalog = "";

            configUser_adminUser = "";
            configUser_adminDomain = "";
            configUser_adminPassword = "";

            configUser_user_Datasource = "";
            configUser_table_view = "";
            configUser_dbTable = "";
            configUser_where = "";
            configUser_user_ad_OU = "";

            configUser_Lname = "";
            configUser_Fname = "";
            configUser_MiddleName = "";
            configUser_StuID = "";
            configUser_password_generate_checkbox = false;
            configUser_password_short_fix_checkbox = false;
            configUser_password = "";
            configUser_levenshtein = false;

            configUser_writeback_AD_checkbox = false;
            configUser_writeback_DB_checkbox = false;
            configUser_writeback_table = "";
            configUser_writeback_primary_key = "";
            configUser_writeback_where_clause = "";
            configUser_writeback_email_field = "";
            configUser_writeback_ad_OU = "";
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

        public String Admin_user
        {
            get
            {
                return configUser_adminUser;
            }
            set
            {
                configUser_adminUser = value;
            }
        }
        public String Admin_domain
        {
            get
            {
                return configUser_adminDomain;
            }
            set
            {
                configUser_adminDomain = value;
            }
        }
        public String Admin_password
        {
            get
            {
                return configUser_adminPassword;
            }
            set
            {
                configUser_adminPassword = value;
            }
        }

        public String User_Datasource
        {
            get
            {
                return configUser_user_Datasource;
            }
            set
            {
                configUser_user_Datasource = value;
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
        public String User_ad_OU
        {
            get
            {
                return configUser_user_ad_OU;
            }
            set
            {
                configUser_user_ad_OU = value;
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
        public String User_Mname
        {
            get
            {
                return configUser_MiddleName;
            }
            set
            {
                configUser_MiddleName = value;
            }
        }
        public String User_StuID
        {
            get
            {
                return configUser_StuID;
            }
            set
            {
                configUser_StuID = value;
            }
        }
        public Boolean User_password_generate_checkbox
        {
            get
            {
                return configUser_password_generate_checkbox;
            }
            set
            {
                configUser_password_generate_checkbox = value;
            }
        }
        public Boolean User_password_short_fix_checkbox
        {
            get
            {
                return configUser_password_short_fix_checkbox;
            }
            set
            {
                configUser_password_short_fix_checkbox = value;
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
        public Boolean Levenshtein
        {
            get
            {
                return configUser_levenshtein;
            }
            set
            {
                configUser_levenshtein = value;
            }
        }

        public Boolean Writeback_AD_checkbox
        {
            get
            {
                return configUser_writeback_AD_checkbox;
            }
            set
            {
                configUser_writeback_AD_checkbox = value;
            }
        }
        public Boolean Writeback_DB_checkbox
        {
            get
            {
                return configUser_writeback_DB_checkbox;
            }
            set
            {
                configUser_writeback_DB_checkbox = value;
            }
        }
        public String Writeback_table
        {
            get
            {
                return configUser_writeback_table;
            }
            set
            {
                configUser_writeback_table = value;
            }
        }
        public String Writeback_primary_key
        {
            get
            {
                return configUser_writeback_primary_key;
            }
            set
            {
                configUser_writeback_primary_key = value;
            }
        }
        public Boolean Writeback_transfer_email_checkbox
        {
            get
            {
                return configUser_writeback_transfer_email_checkbox;
            }
            set
            {
                configUser_writeback_transfer_email_checkbox = value;
            }
        }
        public String Writeback_where_clause
        {
            get
            {
                return configUser_writeback_where_clause;
            }
            set
            {
                configUser_writeback_where_clause = value;
            }
        }
        public String Writeback_email_field
        {
            get
            {
                return configUser_writeback_email_field;
            }
            set
            {
                configUser_writeback_email_field = value;
            }
        }
        public String Writeback_secondary_email_field
        {
            get
            {
                return configUser_writeback_secondary_email_field;
            }
            set
            {
                configUser_writeback_secondary_email_field = value;
            }
        }
        public String Writeback_ad_OU
        {
            get
            {
                return configUser_writeback_ad_OU;
            }
            set
            {
                configUser_writeback_ad_OU = value;
            }
        }


        // creates a dictionay of values
        public Dictionary<string, string> ToDictionary()
        {
            Dictionary<string, string> returnvalue = new Dictionary<string, string>();
            returnvalue.Add("configDataServer", configDataServer);
            returnvalue.Add("configDBCatalog", configDBCatalog);

            returnvalue.Add("configUser_adminUser", configUser_adminUser);
            returnvalue.Add("configUser_adminDomain", configUser_adminDomain);
            returnvalue.Add("configUser_adminPassword", configUser_adminPassword);

            returnvalue.Add("configUser_user_Datasource", configUser_user_Datasource);
            returnvalue.Add("configUser_table_view", configUser_table_view);
            returnvalue.Add("configUser_dbTable", configUser_dbTable);
            returnvalue.Add("configUser_where", configUser_where);
            returnvalue.Add("configUser_user_ad_OU", configUser_user_ad_OU);

            returnvalue.Add("configUser_Lname", configUser_Lname);
            returnvalue.Add("configUser_Fname", configUser_Fname);
            returnvalue.Add("configUser_MiddleName", configUser_MiddleName);
            returnvalue.Add("configUser_StuID", configUser_StuID);
            returnvalue.Add("configUser_password_generate_checkbox", configUser_password_generate_checkbox.ToString());
            returnvalue.Add("configUser_password_short_fix_checkbox", configUser_password_short_fix_checkbox.ToString());
            returnvalue.Add("configUser_password", configUser_password);
            returnvalue.Add("configUser_levenshtein", configUser_writeback_AD_checkbox.ToString());

            returnvalue.Add("configUser_writeback_AD_checkbox", configUser_writeback_AD_checkbox.ToString());
            returnvalue.Add("configUser_writeback_DB_checkbox", configUser_writeback_DB_checkbox.ToString());
            returnvalue.Add("configUser_writeback_table", configUser_writeback_table);
            returnvalue.Add("configUser_writeback_primary_key", configUser_writeback_primary_key);
            returnvalue.Add("configUser_writeback_transfer_email_checkbox", configUser_writeback_transfer_email_checkbox.ToString());
            returnvalue.Add("configUser_writeback_where_clause", configUser_writeback_where_clause);
            returnvalue.Add("configUser_writeback_email_field", configUser_writeback_email_field);
            returnvalue.Add("configUser_writeback_secondary_email_field", configUser_writeback_secondary_email_field);
            returnvalue.Add("configUser_writeback_ad_OU", configUser_writeback_ad_OU);

            return returnvalue;
        }
        public void Load(Dictionary<string, string> dictionary)
        {
            dictionary.TryGetValue("configDataServer", out configDataServer);
            dictionary.TryGetValue("configDBCatalog", out configDBCatalog);

            dictionary.TryGetValue("configUser_adminUser", out configUser_adminUser);
            dictionary.TryGetValue("configUser_adminDomain", out configUser_adminDomain);
            dictionary.TryGetValue("configUser_adminPassword", out configUser_adminPassword);

            dictionary.TryGetValue("configUser_user_Datasource", out configUser_user_Datasource);
            dictionary.TryGetValue("configUser_table_view", out configUser_table_view);
            dictionary.TryGetValue("configUser_dbTable", out configUser_dbTable);
            dictionary.TryGetValue("configUser_where", out configUser_where);
            dictionary.TryGetValue("configUser_user_ad_OU", out configUser_user_ad_OU);

            dictionary.TryGetValue("configUser_Lname", out configUser_Lname);
            dictionary.TryGetValue("configUser_Fname", out configUser_Fname);
            dictionary.TryGetValue("configUser_MiddleName", out configUser_MiddleName);
            dictionary.TryGetValue("configUser_StuID", out configUser_StuID);
            dictionary.TryGetValue("configUser_password_generate_checkbox", out boolconv);
            if (boolconv == "True")
            {
                configUser_password_generate_checkbox = true;
            }
            else
            {
                configUser_password_generate_checkbox = false;
            }
            dictionary.TryGetValue("configUser_password_short_fix_checkbox", out boolconv);
            if (boolconv == "True")
            {
                configUser_password_short_fix_checkbox = true;
            }
            else
            {
                configUser_password_short_fix_checkbox = false;
            }
            dictionary.TryGetValue("configUser_password", out configUser_password);
            dictionary.TryGetValue("configUser_levenshtein", out boolconv);
            if (boolconv == "True")
            {
                configUser_levenshtein = true;
            }
            else
            {
                configUser_levenshtein = false;
            }

            dictionary.TryGetValue("configUser_writeback_AD_checkbox", out boolconv);
            if (boolconv == "True")
            {
                configUser_writeback_AD_checkbox = true;
            }
            else
            {
                configUser_writeback_AD_checkbox = false;
            }
            dictionary.TryGetValue("configUser_writeback_DB_checkbox", out boolconv);
            if (boolconv == "True")
            {
                configUser_writeback_DB_checkbox = true;
            }
            else
            {
                configUser_writeback_DB_checkbox = false;
            }
            dictionary.TryGetValue("configUser_writeback_table", out configUser_writeback_table);
            dictionary.TryGetValue("configUser_writeback_primary_key", out configUser_writeback_primary_key);
            dictionary.TryGetValue("configUser_writeback_transfer_email_checkbox", out boolconv);
            if (boolconv == "True")
            {
                configUser_writeback_transfer_email_checkbox = true;
            }
            else
            {
                configUser_writeback_transfer_email_checkbox = false;
            }
            dictionary.TryGetValue("configUser_writeback_where_clause", out configUser_writeback_where_clause);
            dictionary.TryGetValue("configUser_writeback_email_field", out configUser_writeback_email_field);
            dictionary.TryGetValue("configUser_writeback_secondary_email_field", out configUser_writeback_secondary_email_field);
            dictionary.TryGetValue("configUser_writeback_ad_OU", out configUser_writeback_ad_OU);

        }

    }
    public class ConfigSettings
    {
        private Boolean configTempTables_checkbox;
        private String configLogType;
        private String configLogDB;
        private String configLogCatalog;
        private String configLogDirectory;
        private string boolconv = "";

        public ConfigSettings()
        {
            configTempTables_checkbox = true;
            configLogType = "";
            configLogDB = "";
            configLogCatalog = "";
            configLogDirectory = "";
        }

        // accessor to properties
        //public String progress
        //{
        //    get
        //    {
        //        return configprogress;
        //    }
        //    set
        //    {
        //        configprogress = value;
        //    }
        //}
        public Boolean TempTables
        {
            get
            {
                return configTempTables_checkbox;
            }
            set
            {
                configTempTables_checkbox = value;
            }
        }
        public String LogType
        {
            get
            {
                return configLogType;
            }
            set
            {
                configLogType = value;
            }
        }
        public String LogDB
        {
            get
            {
                return configLogDB;
            }
            set
            {
                configLogDB = value;
            }
        }
        public String LogCatalog
        {
            get
            {
                return configLogCatalog;
            }
            set
            {
                configLogCatalog = value;
            }
        }
        public String LogDirectory
        {
            get
            {
                return configLogDirectory;
            }
            set
            {
                configLogDirectory = value;
            }
        }

        // creates a dictionay of values
        public Dictionary<string, string> ToDictionary()
        {
            Dictionary<string, string> returnvalue = new Dictionary<string, string>();
            returnvalue.Add("configTempTables", configTempTables_checkbox.ToString());
            returnvalue.Add("configLogType", (configLogType != null) ? configLogType.ToString() : "");
            returnvalue.Add("configLogDB", (configLogDB != null) ? configLogDB.ToString() : "");
            returnvalue.Add("configLogCatalog", (configLogCatalog != null) ? configLogCatalog.ToString() : "");
            returnvalue.Add("configLogDirectory", (configLogDirectory != null) ? configLogDirectory.ToString() : "") ;
            return returnvalue;
        }
        public void Load(Dictionary<string, string> dictionary)
        {
            dictionary.TryGetValue("configTempTables", out boolconv);
            if (boolconv == "True")
            {
                configTempTables_checkbox = true;
            }
            else
            {
                configTempTables_checkbox = false;
            }
            dictionary.TryGetValue("configLogType", out configLogType);
            dictionary.TryGetValue("configLogDB", out configLogDB);
            dictionary.TryGetValue("configLogCatalog", out configLogCatalog);
            dictionary.TryGetValue("configLogDirectory", out configLogDirectory);
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

    // tools for working with EVERYTHING
    public class ToolSet
    {

        // AD Functions
        public string GetDomain()
        {
            try
            {
                using (Domain d = Domain.GetCurrentDomain())
                using (DirectoryEntry entry = d.GetDirectoryEntry())
                {
                    return entry.Path;
                }
            }
            catch
            {
                return "fabrikam.com";
            }
        }
        public DataTable EnumerateUsersInOUDataTable(string ouDN, ArrayList returnProperties, string table, SearchScope scope, LogFile log)
        {
            // note does not handle special/illegal characters for AD
            // RETURNS ALL USERS IN AN OU NO INCULDING SUBLEVELS MATTER HOW DEEP 
            int count = returnProperties.Count;
            int i;
            DataTable returnvalue = new DataTable();
            DataRow row;


            // bind to the OU you want to enumerate
            try
            {
                DirectoryEntry deOU = new DirectoryEntry("LDAP://" + ouDN);

                // create a directory searcher for that OU
                DirectorySearcher dsUsers = new DirectorySearcher(deOU);

                // set depth to recursive
                dsUsers.SearchScope = scope;

                // set the filter to get just the users
                dsUsers.Filter = "(&(objectClass=user)(objectCategory=Person))";

                // add the attributes you want to grab from the search
                for (i = 0; i < count; i++)
                {
                    dsUsers.PropertiesToLoad.Add(returnProperties[i].ToString());
                    returnvalue.Columns.Add(returnProperties[i].ToString());
                }
                //dsUsers.PropertiesToLoad.Add("sAMAccountName");

                // grab the users and do whatever you need to do with them 
                dsUsers.PageSize = 1000;
                row = returnvalue.NewRow();

                foreach (SearchResult oResult in dsUsers.FindAll())
                {
                    //generate the array list with the user sam accounts
                    for (i = 0; i < count; i++)
                    {
                        try
                        {
                            row[i] = System.Web.HttpUtility.UrlDecode((Convert.IsDBNull(oResult.Properties[returnProperties[i].ToString()][0]) ? string.Empty : oResult.Properties[returnProperties[i].ToString()][0].ToString()));
                        }
                        catch
                        {
                            row[i] = string.Empty;
                        }
                    }
                    returnvalue.Rows.Add(row);
                    row = returnvalue.NewRow();
                }
                dsUsers.Dispose();
            }

            catch (Exception ex)
            {
                log.addTrn("Failure getting AD users exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }

            return returnvalue;
        }
        public DataTable EnumerateUsersInGroupDataTable(string groupDN, string groupou, string coulumnNameForFQDN, string coulumnNameForGroup, string table, LogFile log)
        {
            // note does not handle special/illegal characters for AD
            // it is optimal to have the two field names match the coumn names pulled from SQL do not use group fro group name it will kill the SQL
            // groupDN "CN=Sales,OU=test,DC=Fabrikam,DC=COM"
            // returns FQDN "CN=user,OU=test,DC=Fabrikam,DC=COM" & group "CN=Sales,OU=test,DC=Fabrikam,DC=COM" of users in group 
            DataTable returnvalue = new DataTable();
            DataRow row;
            int count = 0;

            //string bladh = "LDAP://" + "CN=" + System.Web.HttpUtility.UrlEncode(groupDN).Replace("+", " ").Replace("*", "%2A") + groupou;

                //DirectoryEntry group = new DirectoryEntry("LDAP://" + "CN=" + System.Web.HttpUtility.UrlEncode(groupDN).Replace("+", " ").Replace("*", "%2A") + groupou);
                //DirectorySearcher groupUsers = new DirectorySearcher(group);

                
                try
                {
                    DirectoryEntry group = new DirectoryEntry("LDAP://" + "CN=" + System.Web.HttpUtility.UrlEncode(groupDN).Replace("+", " ").Replace("*", "%2A") + groupou);
                    DirectorySearcher groupUsers = new DirectorySearcher(group);
                    groupUsers.Filter = "(objectClass=*)";
                    row = returnvalue.NewRow();
                    returnvalue.TableName = table;
                    returnvalue.Columns.Add(coulumnNameForFQDN);
                    returnvalue.Columns.Add(coulumnNameForGroup);
                    uint rangeStep = 1000;
                    uint rangeLow = 0;
                    uint rangeHigh = rangeLow + (rangeStep - 1);
                    bool lastQuery = false;
                    bool quitLoop = false;

                    do
                    {
                        string attributeWithRange;
                        if (!lastQuery)
                        {
                            attributeWithRange = String.Format("member;range={0}-{1}", rangeLow, rangeHigh);
                        }
                        else
                        {
                            attributeWithRange = String.Format("member;range={0}-*", rangeLow);
                        }
                        groupUsers.PropertiesToLoad.Clear();
                        groupUsers.PropertiesToLoad.Add(attributeWithRange);
                        SearchResult results = groupUsers.FindOne();
                        groupUsers.Dispose();
                        foreach (string res in results.Properties.PropertyNames)
                        {
                            System.Diagnostics.Debug.WriteLine(res.ToString());
                        }
                        if (results.Properties.Contains(attributeWithRange))
                        {
                            foreach (object obj in results.Properties[attributeWithRange])
                            {
                                Console.WriteLine(obj.GetType());
                                if (obj.GetType().Equals(typeof(System.String)))
                                {
                                }
                                else if (obj.GetType().Equals(typeof(System.Int32)))
                                {
                                }
                                row[0] = obj.ToString();
                                row[1] = groupDN;
                                returnvalue.Rows.Add(row);
                                row = returnvalue.NewRow();
                                count++;
                            }
                            if (lastQuery)
                            {
                                quitLoop = true;
                            }
                        }
                        else
                        {
                            lastQuery = true;
                        }
                        if (!lastQuery)
                        {
                            rangeLow = rangeHigh + 1;
                            rangeHigh = rangeLow + (rangeStep - 1);
                        }
                        // if we are searching for the next set of members 1000-* and it did not return any records count == 0
                        if (attributeWithRange == String.Format("member;range={0}-*", rangeLow) && count == 0)
                        {
                            quitLoop = true;
                        }
                        count = 0;
                    }
                    while (!quitLoop);
                }
            catch (Exception ex)
            {
                log.addTrn("Failure getting AD users in a groups exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
            //foreach (object dn in group.properties["member"])
            //{
            //    row[0] = dn.tostring();
            //    row[1] = groupDN;
            //    returnvalue.row.add(row);
            //    row = returnvalue.newrow();

            //}
            

            return returnvalue;
        }
        public DataTable EnumerateGroupsInOUDataTable(string ouDN, ArrayList returnProperties, string table, LogFile log)
        {
            // note does not handle special/illegal characters for AD
            int i;
            int count = returnProperties.Count;
            DataTable returnvalue = new DataTable();
            DataRow row;
            // bind to the OU you want to enumerate
            try
            {
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
                dsUsers.PageSize = 1000;
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
                dsUsers.Dispose();
            }
            catch (Exception ex)
            {
                log.addTrn("Failure getting AD groups exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
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
                catch
                {
                    // MessageBox.Show(e.Message.ToString() + "create group LDAP://CN=" + name + "," + ouPath);
                    return found;
                }
            }
            return found;
        }

        public void SetAttributeValuesSingleString(string attributeName, string newValue, string objectDn, LogFile log)
        {
            // objectDn expects full ldap entery LDAP://cn=blah,ou=testing,dc=fabrikam,dc=com

            try
            {
                DirectoryEntry ent = new DirectoryEntry(objectDn);
                ent.Properties[attributeName].Value = newValue;
                ent.CommitChanges();
                log.addTrn("AD set value for field " + attributeName + " for user " + objectDn + " value " + newValue, "Transaction");
                ent.Close();
                ent.Dispose();
            }
            catch (Exception ex)
            {
                log.addTrn("failed AD set value for field " + attributeName + " for user " + objectDn + " value " + newValue + " exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }


        }
        public string GetObjectDistinguishedName(objectClass objectCls, returnType returnValue, string objectName, string ldapDomain, LogFile log)
        {
            // LdapDomain = "DC=Fabrikam,DC=COM" 

            string distinguishedName = string.Empty;
            string connectionPrefix = "LDAP://" + ldapDomain;
            try
            {
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
                    case objectClass.organizationalunit:
                        mySearcher.Filter = "(&(objectClass=organizationalunit)(distinguishedname=" + objectName + "))";
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
            }
            catch (Exception ex)
            {
                log.addTrn("searcher failed " + ldapDomain + " " + objectName + " Exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }

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
                log.addTrn(E.Message.ToString() + " error creating OU" + ou, "Error");
                return false;
            }
        }
        public void CreateGroup(string ouPath, Dictionary<string, string> properties, LogFile log)
        {
            // otherProperties is a mapping  <the key is the active directory field, and the value is the the value>
            // the keys must contain valid AD fields
            // the value will relate to the specific key
            // needs parent OU present to work
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
                    log.addTrn("group added | LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath, "Transaction");
                }
                else
                {
                    log.addTrn("CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + " group already exists from adding", "Warning");
                }
            }
            catch (Exception ex)
            {
                log.addTrn(ex.Message.ToString() + "issue create group LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + "\n" + ex.StackTrace.ToString(), "Error");
            }
        }
        public void UpdateGroup(string ouPath, Dictionary<string, string> properties, LogFile log)
        {
            // otherProperties is a mapping  <the key is the active driectory field, and the value is the the value>
            // the keys must contain valid AD fields
            // the value will relate to the specific key
            // needs parent OU present to work
            try
            {
                if (DirectoryEntry.Exists("LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath))
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
                    log.addTrn("updated group | LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath, "Transaction");

                }
                else
                {
                    log.addTrn(ouPath + " group does not exist", "Warning");
                }
            }
            catch (Exception ex)
            {
                log.addTrn(ex.Message.ToString() + "issue updating group LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + "\n" + ex.StackTrace.ToString(), "Error");
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
                    log.addTrn("deleted group | LDAP://CN=" + name + "," + ouPath, "Transaction");
                }
                catch (Exception ex)
                {
                    log.addTrn(ex.Message.ToString() + " error deleting LDAP://CN=" + name + "," + ouPath + "\n" + ex.StackTrace.ToString(), "Error");
                }
            }
            else
            {
                log.addTrn("group LDAP://CN=" + name + "," + ouPath + " does not exists cannot delete", "Warning");
            }
        }
        public void CreateOU(string ouPath, string name, LogFile log)
        {
            //needs parent OU present to work
            try
            {
                if (!DirectoryEntry.Exists("LDAP://OU=" + name + "," + ouPath))
                {

                    DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                    DirectoryEntry OU = entry.Children.Add("OU=" + name, "organizationalUnit");
                    OU.CommitChanges();
                    OU.Close();
                    OU.Dispose();
                    entry.Close();
                    entry.Dispose();
                    log.addTrn("created ou | LDAP://OU=" + name + "," + ouPath, "Transaction");

                }
                else
                {
                    log.addTrn("creating ou LDAP://OU=" + name + "," + ouPath + " already exists", "Warning");
                }
            }
            catch (Exception ex)
            {
                log.addTrn(ex.Message.ToString() + "error creating ou LDAP://OU=" + name + "," + ouPath + "\n" + ex.StackTrace.ToString(), "Error");
            }
        }
        public void DeleteOU(string ouPath, string name, LogFile log)
        {
            try
            {
                //needs parent OU present to work
                if (!DirectoryEntry.Exists("LDAP://OU=" + name + "," + ouPath))
                {

                    DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                    entry.DeleteTree();
                    entry.Close();
                    entry.Dispose();
                    log.addTrn("deleting ou | LDAP://OU=" + name + "," + ouPath + " does not exists", "Transaction");

                }
                else
                {
                    log.addTrn("error deleting ou LDAP://OU=" + name + "," + ouPath + " does not exists", "Warning");
                }
            }
            catch (Exception ex)
            {
                log.addTrn(ex.Message.ToString() + " error deleting ou LDAP://OU=" + name + "," + ouPath + "\n" + ex.StackTrace.ToString(), "Error");
            }
        }
        public void AddUserToGroup(string userDn, string groupDn, bool search, string ldapDomain, LogFile log)
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
                    log.addTrn("added user to group | " + userDn + " | LDAP://" + groupDn, "Transaction");
                }
                else
                {
                    if (search)
                    {

                        DirectoryEntry entry = new DirectoryEntry("LDAP://" + groupDn);
                        userDn = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, userDn.Substring((userDn.IndexOf("CN=") + 3), (userDn.IndexOf(",OU") - 3)), ldapDomain, log);
                        entry.Properties["member"].Add(userDn);
                        // dirEntry.Invoke("Add", new object[] { "LDAP://" + userDn });
                        entry.CommitChanges();
                        entry.Close();
                        entry.Dispose();
                        log.addTrn("Had to find user and then added user to group | " + userDn + " | LDAP://" + groupDn, "Transaction");
                    }
                    else
                    {
                        log.addTrn(" Warning could not add user " + userDn + " to group LDAP://" + groupDn + " group did not exist", "Warning");
                    }

                }

            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                log.addTrn(E.Message.ToString() + " error adding user to group" + userDn + " to LDAP://" + groupDn, "Error");

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
                    log.addTrn("removed user from group | " + userDn + " | LDAP://" + groupDn, "Transaction");
                }
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    log.addTrn(E.Message.ToString() + " error removing user from group " + userDn + " from LDAP://" + groupDn + " user may not be in group", "Error");
                }
                entry.CommitChanges();
                entry.Close();
                entry.Dispose();
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                log.addTrn(E.Message.ToString() + " error removing user from group " + userDn + " from LDAP://" + groupDn + " group object does not exist", "Warning");
            }
        }


        public void CreateUsersAccounts(string ouPath, SqlDataReader users, string groupDn, string ldapDomain, UserSynch usersyn, LogFile log)
        {
            // oupath holds the path for the AD OU to hold the Users 
            // users is a sqldatareader witht the required fields in it ("CN") other Datastructures would be easy to substitute 
            // groupDN is a base group which all new users get automatically inserted into

            int i;
            int fieldcount;
            int val;
            string name = "";
            string last = "";
            string first = "";
            fieldcount = users.FieldCount;
            try
            {
                while (users.Read())
                {
                    try
                    {

                        if (users[usersyn.User_password].ToString() != "")
                        {
                            if (!DirectoryEntry.Exists("LDAP://CN=" + System.Web.HttpUtility.UrlEncode(users[usersyn.User_sAMAccount].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath))
                            {

                                DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                                DirectoryEntry newUser = entry.Children.Add("CN=" + System.Web.HttpUtility.UrlEncode(users[usersyn.User_CN].ToString()).Replace("+", " ").Replace("*", "%2A"), "user");
                                // generated
                                newUser.Properties["samAccountName"].Value = System.Web.HttpUtility.UrlEncode(users[usersyn.User_sAMAccount].ToString()).Replace("+", " ").Replace("*", "%2A");
                                //newUser.Properties["mail"].Value = System.Web.HttpUtility.UrlEncode(users[usersyn.User_mail].ToString()).Replace("+", " ").Replace("*", "%2A") + "@" + System.Web.HttpUtility.UrlEncode(users[usersyn.UserEmailDomain].ToString()).Replace("+", " ").Replace("*", "%2A");
                                newUser.Properties["UserPrincipalName"].Value = System.Web.HttpUtility.UrlEncode(users[usersyn.User_sAMAccount].ToString()).Replace("+", " ").Replace("*", "%2A");
                                newUser.Properties["displayName"].Value = System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Lname]).Replace("+", " ").Replace("*", "%2A") + ", " + System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Fname]).Replace("+", " ").Replace("*", "%2A");
                                newUser.Properties["description"].Value = System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Lname]).Replace("+", " ").Replace("*", "%2A") + ", " + System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Fname]).Replace("+", " ").Replace("*", "%2A");

                                newUser.CommitChanges();

                                // SQL query generated ensures matching field names between the SQL form fields and AD
                                for (i = 0; i < fieldcount; i++)
                                {
                                    name = users.GetName(i);
                                    // eliminiate non updatable fields
                                    if (name != "password" && name != "CN")
                                    {
                                        // mail needs some special handling
                                        if (name != "mail")
                                        {
                                            if ((string)users[name] != "")
                                            {
                                                newUser.Properties[name].Value = System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A");
                                            }
                                        }
                                        else
                                        {
                                            first = System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A").Replace("%40", "@");
                                            last = (string)users[name];
                                            // check to see if mail field has illegal characters
                                            if (first == last)
                                            {
                                                // no illegal characters input the value into AD
                                                newUser.Properties[name].Value = System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A").Replace("!", "%21").Replace("(", "%28").Replace(")", "%29").Replace("'", "%27").Replace("_", "%5f").Replace(" ", "%20").Replace("%40", "@");
                                            }
                                            else
                                            {
                                                // newUser.Properties[name].Value = "";
                                            }
                                        }
                                    }
                                }


                                AddUserToGroup("CN=" + System.Web.HttpUtility.UrlEncode(users[usersyn.User_sAMAccount].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + usersyn.UserHoldingTank, groupDn, false, ldapDomain, log);
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
                                log.addTrn("User added |" + (string)users[usersyn.User_sAMAccount] + " " + usersyn.UserHoldingTank, "Transaction");
                            }
                            else
                            {
                                log.addTrn("CN=" + System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_sAMAccount]).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + " user already exists from adding", "Error");
                                //MessageBox.Show("CN=" + System.Web.HttpUtility.UrlEncode((string)users["CN"]).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + " user already exists from adding");
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        string debugdata = "";
                        for (i = 0; i < fieldcount; i++)
                        {

                            debugdata += users.GetName(i) + "=" + System.Web.HttpUtility.UrlEncode((string)users[i]).Replace("+", " ").Replace("*", "%2A") + ", ";

                        }
                        log.addTrn("issue create user LDAP://CN=" + System.Web.HttpUtility.UrlEncode((string)users["CN"]).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + "\n" + debugdata + " User create failed, commit error" + name + " | " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                        // MessageBox.Show(e.Message.ToString() + "issue create user LDAP://CN=" + System.Web.HttpUtility.UrlEncode((string)users["CN"]).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + "\n" + debugdata);
                    }
                }

            }
            catch (Exception ex)
            {
                if (users != null)
                {
                    string debugdata = "";
                    for (i = 0; i < fieldcount; i++)
                    {

                        debugdata += users.GetName(i) + "=" + System.Web.HttpUtility.UrlEncode((string)users[i]).Replace("+", " ").Replace("*", "%2A") + ", ";

                    }
                    log.addTrn("issue create user LDAP://CN=" + System.Web.HttpUtility.UrlEncode((string)users["CN"]).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + "\n" + debugdata + " failed field maybe " + name + " | " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                    // MessageBox.Show(e.Message.ToString() + "issue create user LDAP://CN=" + System.Web.HttpUtility.UrlEncode((string)users["CN"]).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + "\n" + debugdata);
                }
                else
                {
                    log.addTrn("issue creating users datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
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
        public void UpdateUsers(SqlDataReader users, string ldapDomain, UserSynch usersyn, LogFile log)
        {
            // requires distinguished name to be a field
            // all field names must be valid AD field names
            // does not blank out fields

            int fieldcount = 0;
            int i = 0;
            string name = "";
            string fdqn = "";
            try
            {
            fieldcount = users.FieldCount;
                while (users.Read())
                {

                    DirectoryEntry user = new DirectoryEntry("LDAP://" + (string)users["distinguishedname"]);
                    for (i = 0; i < fieldcount; i++)
                    {
                        name = users.GetName(i);
                        // eliminiate non updatable fields
                        if (name != "password" && name != "CN" && name != "sAMAccountName" && name != "distinguishedname")
                        {
                            
                            // mail needs some special handling
                            switch (name)
                            {
                                case "mail":
                                    if ((string)users[name] != "")
                                    {
                                        // check to see if mail field has illegal characters
                                        string hi = (System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A").Replace("%40", "@").Replace("%5f", "_"));
                                        string hi3 = (string)users[name];
                                        if (System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A").Replace("%40", "@").Replace("%5f", "_") == (string)users[name])
                                        {
                                            // no illegal characters input the value into AD
                                            user.Properties[name].Value = System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A").Replace("!", "%21").Replace("(", "%28").Replace(")", "%29").Replace("'", "%27").Replace("%5f", "_").Replace(" ", "%20").Replace("%40", "@");
                                        }
                                        else
                                        {
                                            user.Properties[name].Value = "illegal Email";
                                        }
                                    }
                                    break;
                                case "userAccountControl":
                                    if ((string)users[name] != "")
                                    {
                                        int val = (int)user.Properties["userAccountControl"].Value;
                                        user.Properties["userAccountControl"].Value = val | Convert.ToInt32(System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A").ToString());
                                    }
                                    break;
                                case "manager":
                                    if ((string)users[name] != "")
                                    {
                                        fdqn = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A"), ldapDomain, log);
                                        if (!fdqn.Contains("CN"))
                                        {
                                            log.addTrn("Issue Updating User: " + (string)users["distinguishedname"] + " Invalid Manager selected. ", "Error");
                                        }
                                        else
                                        {
                                            user.Properties["manager"].Value = fdqn.Substring(fdqn.IndexOf("CN"));
                                        }
                                    }
                                    break;
                                case "sn":
                                    if ((string)users[name] != "")
                                    {
                                        user.Properties[name].Value = System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A");
                                        user.Properties["displayName"].Value = System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Lname]).Replace("+", " ").Replace("*", "%2A") + ", " + System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Fname]).Replace("+", " ").Replace("*", "%2A");
                                        user.Properties["description"].Value = System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Lname]).Replace("+", " ").Replace("*", "%2A") + ", " + System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Fname]).Replace("+", " ").Replace("*", "%2A");

                                    }
                                    break;
                                case "givenName":
                                    if ((string)users[name] != "")
                                    {
                                        user.Properties[name].Value = System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A");
                                        user.Properties["displayName"].Value = System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Lname]).Replace("+", " ").Replace("*", "%2A") + ", " + System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Fname]).Replace("+", " ").Replace("*", "%2A");
                                        user.Properties["description"].Value = System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Lname]).Replace("+", " ").Replace("*", "%2A") + ", " + System.Web.HttpUtility.UrlEncode((string)users[usersyn.User_Fname]).Replace("+", " ").Replace("*", "%2A");

                                    }
                                    break;
                                default:
                                    if ((string)users[name] != "")
                                    {
                                        user.Properties[name].Value = System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A");
                                    }
                                    break;
                            }


                            //if (name != "mail")
                            //{
                            //    if ((string)users[name] != "")
                            //    {
                            //        user.Properties[name].Value = System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A");
                            //    }
                            //}
                            //else
                            //{
                            //    // check to see if mail field has illegal characters
                            //    string hi = (System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A").Replace("%40", "@"));
                            //    string hi3 = (string)users[name];
                            //    if (System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A").Replace("%40", "@") != (string)users[name])
                            //    {
                            //        // no illegal characters input the value into AD
                            //        user.Properties[name].Value = System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A").Replace("!", "%21").Replace("(", "%28").Replace(")", "%29").Replace("'", "%27").Replace("_", "%5f").Replace(" ", "%20").Replace("%40", "@");
                            //    }
                            //    else
                            //    {
                            //        user.Properties[name].Value = "illegal Email";
                            //    }
                            //}

                        }
                    }
                    user.CommitChanges();
                    log.addTrn("User updated |" + (string)users["distinguishedname"] + " ", "Transaction");
                }
            }
            catch (Exception ex)
            {
                if (users != null)
                {
                    log.addTrn("issue updating user " + name + " " + System.Web.HttpUtility.UrlEncode((string)users["distinguishedname"]).Replace("+", " ").Replace("*", "%2A") + "\n" + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                }
                else
                {
                    log.addTrn("issue updating users data reader is null " + "\n" + ex.Message.ToString(), "Error");
                }
            }
        }
        public bool DisableUser(string sAMAccountName, string ldapDomain, LogFile log)
        {
            string userDN;
            userDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, ldapDomain, log);
            try
            {
                DirectoryEntry usr = new DirectoryEntry(userDN);
                int val = (int)usr.Properties["userAccountControl"].Value;
                usr.Properties["userAccountControl"].Value = val | (int)accountFlags.ADS_UF_ACCOUNTDISABLE;
                usr.CommitChanges();
                usr.Close();
                usr.Dispose();
                log.addTrn("diabled user account |" + userDN, "Transaction");
                return true;
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                log.addTrn(E.Message.ToString() + " error disabling user " + userDN, "Error");
                return false;
            }

        }
        public bool EnableUser(string sAMAccountName, string ldapDomain, LogFile log)
        {
            string userDN;
            userDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, ldapDomain, log);
            try
            {
                DirectoryEntry usr = new DirectoryEntry(userDN);
                int val = (int)usr.Properties["userAccountControl"].Value;
                usr.Properties["userAccountControl"].Value = val | ~(int)accountFlags.ADS_UF_ACCOUNTDISABLE;
                usr.CommitChanges();
                usr.Close();
                usr.Dispose();
                log.addTrn("enabled user account |" + userDN, "Transaction");
                return true;
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                log.addTrn(E.Message.ToString() + " error enabling user " + userDN, "Error");
                return false;
            }
        }
        // additional stuff
        public bool SetUserExpiration(int days, string ldapDomain, string sAMAccountName, LogFile log)
        {
            string userDN;
            userDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, ldapDomain, log);
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
                log.addTrn("User expiration set |" + userDN + "|" + days, "Transaction");

                return true;
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                log.addTrn(E.Message.ToString() + " error setting user expiration " + userDN + " days " + days, "Error");
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
                log.addTrn("User expiration set |" + userDN + "|" + days, "Transaction");
                return true;
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                log.addTrn(E.Message.ToString() + " error setting user expiration " + userDN + " days " + days, "Error");
                return false;
            }
        }
        public bool DeleteUserAccount(string sAMAccountName, string ldapDomain, LogFile log)
        {
            string userDN;
            userDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, ldapDomain, log);
            try
            {
                DirectoryEntry ent = new DirectoryEntry(userDN);
                ent.DeleteTree();
                ent.Close();
                ent.Dispose();
                log.addTrn("deleted user account |" + userDN, "Transaction");
                return true;
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                log.addTrn(E.Message.ToString() + " error deleting user " + userDN, "Error");
                return false;
            }
        }
        public bool DeleteUserAccount(string FQDN, LogFile log)
        {
            try
            {
                FQDN = "LDAP://" + FQDN;
                DirectoryEntry ent = new DirectoryEntry(FQDN);
                ent.DeleteTree();
                ent.Close();
                ent.Dispose();
                log.addTrn("deleted user account |" + FQDN, "Transaction");
                return true;
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                log.addTrn(E.Message.ToString() + " error deleting user " + FQDN, "Error");
                return false;
            }
        }
        public ArrayList ADobjectAttribute()
        {
            // returns a list of all the AD fields in the schema for a user
            // modification to any active directory object type would be simple by changing the directory entry
            // NOTE: One place where managed ADSI (System.DirectoryServices) falls short is finding schema 
            // information from LDAP/AD objects. Finding information like mandatory and optional
            // properties simply cannot be done with any managed classes

            DirectoryEntry schemaEntry = null;
            ArrayList returnvalue = new ArrayList();
            try
            {
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
            }
            catch
            {
                returnvalue.Add("Error pulling AD columns");
            }
            return returnvalue;

        }

        // SQL table tools
        public string Create_Table(DataTable data, string table, SqlConnection sqlConn, LogFile log)
        {
            // sqlConn must be an open connection
            int i;
            int Count;
            StringBuilder sqlstring = new StringBuilder();
            SqlCommand sqlComm;
            Count = data.Columns.Count;

            // make the temp table
            sqlstring.Append("CREATE TABLE " + table + "(");
            for (i = 0; i < Count; i++)
            {
                sqlstring.Append(data.Columns[i] + " VarChar(350), ");
            }
            sqlstring.Remove((sqlstring.Length - 2), 2);
            sqlstring.Append(")");
            sqlComm = new SqlCommand(sqlstring.ToString(), sqlConn);
            try
            {
                sqlComm.CommandTimeout = 360;
                sqlComm.ExecuteNonQuery();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
                log.addTrn("table created " + table, "Transaction");
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }

            // copy data into table
            try
            {
                SqlBulkCopy sbc = new SqlBulkCopy(sqlConn);
                sbc.DestinationTableName = table;
                sbc.WriteToServer(data);

                sbc.Close();
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL bulk copy " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
            return table;
        }
        public string Append_to_Table(DataTable data, string table, SqlConnection sqlConn, LogFile log)
        {
            // copy data into table
            try
            {
                SqlBulkCopy sbc = new SqlBulkCopy(sqlConn);
                sbc.DestinationTableName = table;
                sbc.WriteToServer(data);
                sbc.Close();
            }
            catch (Exception ex)
            {
                log.addTrn("failed sql table append " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                return "failure";
            }
            return table;
        }
        public void DropTable(string table, SqlConnection sqlConn, LogFile log)
        {
            SqlCommand sqlComm;
            string sqlstring = "DROP TABLE " + table;
            sqlComm = new SqlCommand(sqlstring.ToString(), sqlConn);
            try
            {
                sqlComm.CommandTimeout = 360;
                sqlComm.ExecuteNonQuery();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
                log.addTrn("table dropped " + table, "Transaction");
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
        }
        // SQL query tools
        public SqlDataReader QueryNotExists(string table1, string table2, SqlConnection sqlConn, string pkey1, string pkey2, LogFile log)
        {
            // Array list of pkeys is for use when the primary key is clustered (multiple columns are required to get a unique identification on the row)
            // finds items in table1 who do not exist in table2 and returns the data fields table 1 for these rows
            //*************************************************************************************************
            //| Table1                      | Table2                    | Returned result
            //*************************************************************************************************
            //| ID            Data          | ID             Data       |                   | Table1.ID     Table1.DATA
            //| 1             a             | 1              a          | NOT RETURNED      |
            //| 2             b             | null           null       | RETURNED          | 2             b
            //| 3             c             | 3              null       | NOT RETURNED      |
            //| 4             d             | 4              e          | NOT RETURNED      |
            //
            // SqlCommand sqlComm = new SqlCommand("Select Table1.* Into #Table3ADTransfer From " + Table1 + " AS Table1, " + Table2 + " AS Table2 Where Table1." + pkey1 + " = Table2." + pkey2 + " And Table2." + pkey2 + " is null", sqlConn);
            //SqlCommand sqlComm = new SqlCommand("SELECT DISTINCT uptoDate.* FROM " + table1 + " uptoDate LEFT OUTER JOIN " + table2 + " outofDate ON outofDate." + pkey2 + " = uptoDate." + pkey1 + " WHERE outofDate." + pkey2 + " IS NULL;", sqlConn);
            SqlCommand sqlComm = new SqlCommand("SELECT * FROM " + table1 + " uptoDate EXCEPT  SELECT * FROM " + table2 + " AS outofDate", sqlConn);
            // create the command object
            SqlDataReader r;
            try
            {
                sqlComm.CommandTimeout = 360;
                r = sqlComm.ExecuteReader();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
                return r;
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
            return null;
        }
        public void QueryNotExistsIntoNewTable(string table1, string table2, string newTable, SqlConnection sqlConn, string pkey1, string pkey2, LogFile log)
        {
            // finds items in table1 who do not exist in table2 and creates a table with the data fields from table 1
            //*************************************************************************************************
            //| Table1                      | Table2                    | Returned result
            //*************************************************************************************************
            //| ID            Data          | ID             Data       |                   | Table1.ID     Table1.DATA
            //| 1             a             | 1              a          | NOT RETURNED      |
            //| 2             b             | null           null       | RETURNED          | 2             b
            //| 3             c             | 3              null       | NOT RETURNED      |
            //| 4             d             | 4              e          | NOT RETURNED      |
            // SqlCommand sqlComm = new SqlCommand("Select Table1.* Into #Table3ADTransfer From " + Table1 + " AS Table1, " + Table2 + " AS Table2 Where Table1." + pkey1 + " = Table2." + pkey2 + " And Table2." + pkey2 + " is null", sqlConn);
            SqlCommand sqlComm = new SqlCommand("SELECT DISTINCT uptoDate.* INTO " + newTable + " FROM " + table1 + " uptoDate LEFT OUTER JOIN " + table2 + " outofDate ON outofDate." + pkey2 + " = uptoDate." + pkey1 + " WHERE outofDate." + pkey2 + " IS NULL", sqlConn);
            // create the command object
            try
            {
                sqlComm.CommandTimeout = 360;
                sqlComm.ExecuteNonQuery();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }


        }
        public SqlDataReader QueryInnerJoin(string table1, string table2, string pkey1, string pkey2, SqlConnection sqlConn, LogFile log)
        {
            // Returns data from table1 where the row is in both table 1 and table2
            //*************************************************************************************************
            //| Table1                      | Table2                    | Returned result
            //*************************************************************************************************
            //| ID            Data          | ID             Data       |                   | Table1.ID     Table1.DATA
            //| 1             a             | 1              a          | RETURNED          | 1             a 
            //| 2             b             | null           null       | NOT RETURNED      |              
            //| 3             c             | 3              null       | RETURNED          | 3             c
            //| 4             d             | 4              e          | RETURNED          | 4             d

            SqlCommand sqlComm = new SqlCommand("SELECT DISTINCT " + table1 + ".* FROM " + table1 + " INNER JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2, sqlConn);
            SqlDataReader r;
            try
            {
                sqlComm.CommandTimeout = 360;
                r = sqlComm.ExecuteReader();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
                return r;
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
            return null;
        }
        public SqlDataReader QueryInnerJoin(string table1, string table2, string pkey1, string pkey2, ArrayList additionalFields, SqlConnection sqlConn, LogFile log)
        {
            // additionalFields takes the field names " table.field,"
            // Returns data from table1 where the row is in both table 1 and table2
            // additional fields table2.data2
            //*************************************************************************************************
            //| Table1                      | Table2                                    | Returned result
            //*************************************************************************************************
            //| ID            Data          | ID             Data         Data2         |                   | Table1.ID     Table1.DATA     Table2.data2
            //| 1             a             | 1              a            e             | RETURNED          | 1             a               e
            //| 2             b             | null           null         f             | NOT RETURNED      |              
            //| 3             c             | 3              null         g             | RETURNED          | 3             c               g
            //| 4             d             | 4              e            h             | RETURNED          | 4             d               h

            SqlDataReader r;
            SqlCommand sqlComm;
            string additionalfields = "";
            foreach (string key in additionalFields)
            {
                additionalfields += key;
            }
            additionalfields = additionalfields.Remove(additionalfields.Length - 2);
            if (additionalFields.Count > 0)
            {
                sqlComm = new SqlCommand("SELECT DISTINCT " + table1 + ".*, " + additionalfields + " FROM " + table1 + " INNER JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2, sqlConn);
            }
            else
            {
                sqlComm = new SqlCommand("SELECT DISTINCT " + table1 + ".* FROM " + table1 + " INNER JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2, sqlConn);
            }

            try
            {
                sqlComm.CommandTimeout = 360;
                r = sqlComm.ExecuteReader();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
                return r;
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
            return null;
        }
        public void QueryInnerJoinIntoNewTable(string table1, string table2, string pkey1, string pkey2, string newTable, SqlConnection sqlConn, LogFile log)
        {
            // Creates a table with data from table1 where the row is in both table 1 and table2
            //*************************************************************************************************
            //| Table1                      | Table2                    | Returned result
            //*************************************************************************************************
            //| ID            Data          | ID             Data       |                   | Table1.ID     Table1.DATA            ----> Table1.*
            //| 1             a             | 1              a          | RETURNED          | 1             a 
            //| 2             b             | null           null       | NOT RETURNED      |              
            //| 3             c             | 3              null       | RETURNED          | 3             c
            //| 4             d             | 4              e          | RETURNED          | 4             d 

            SqlCommand sqlComm = new SqlCommand("SELECT DISTINCT " + table1 + ".* INTO " + newTable + " FROM " + table1 + " INNER JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2, sqlConn);
            try
            {
                sqlComm.CommandTimeout = 360;
                sqlComm.ExecuteNonQuery();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
        }
        public SqlDataReader CheckUpdate(string table1, string table2, string pkey1, string pkey2, ArrayList compareFields1, ArrayList compareFields2, SqlConnection sqlConn, LogFile log)
        {
            // NULL not handled as blanks
            // Assumes table1 holds the correct data and returns a data reader with the update fields columns from table1
            // returns the rows which table2's concatenated update fields differ from table1's concatenated update fields
            // eliminates rows which do not have a matching key in both tables
            //*************************************************************************************************
            //| Table1                      | Table2                    | Returned result
            //*************************************************************************************************
            //| ID            Data          | ID             Data       |                   | TABLE1.ID     TABLE1.DATA
            //| 1             a             | 1              a          | NOT RETURNED      |
            //| 2             b             | null           null       | NOT RETURNED      |              
            //| 3             c             | 3              null       | RETURNED          | 3             c
            //| 4             d             | 4              e          | RETURNED          | 4             d
            //| 4             f             | 4              f          | NOT RETURNED      | 

            string compare1 = "";
            string compare2 = "";
            string fields1 = "";
            string fields2 = "";
            // need a comand builder and research on the best way to compare all fields in a row
            // this basically will just issue a concatenation sql query to the DB for each field to compare
            foreach (string key in compareFields1)
            {
                fields1 += table1 + "." + key + ", ";
            }
            foreach (string key in compareFields2)
            {
                fields2 += table2 + "." + key + ", ";
            }
            
            // remove trailing comma and + 
            compare2 = compare2.Remove(compare2.Length - 2);
            compare1 = compare1.Remove(compare1.Length - 2);
            fields1 = fields1.Remove(fields1.Length - 2);
            fields2 = fields2.Remove(fields2.Length - 2);

            SqlCommand sqlComm = new SqlCommand(    "SELECT " + fields1 +
                                                    " FROM " + table1 +
                                                    " INNER JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2 +
                                                    " GROUP BY " + fields1 +
                                                    " EXCEPT " +
                                                    " SELECT " + fields2 + " FROM " + table2, sqlConn);

            
      
            //SqlCommand sqlComm = new SqlCommand("SELECT DISTINCT " + fields1 + " FROM " + table1 + " LEFT JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2 + " AND (" + compare2 + ") <> (" + compare1 + ") INNER JOIN " + table2 + " as [" + table2 + "temp] ON " + table1 + "." + pkey1 + " = [" + table2 + "temp]." + pkey2 + " WHERE " + table2 + "." + pkey2 + " IS NOT NULL", sqlConn);
            
            try
            {
                sqlComm.CommandTimeout = 360;
                SqlDataReader r = sqlComm.ExecuteReader();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
                return r;
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
            return null;
        }
        public SqlDataReader CheckUpdate(string table1, string table2, string pkey1, string pkey2, ArrayList compareFields1, ArrayList compareFields2, ArrayList additionalFields, int adField, SqlConnection sqlConn, LogFile log)
        {
            // adField holds the value of which table holds values from ad to check the AD field fomat distinguished name vs samaccount / id if ad field is 0 neither should be checked
            // managerADtype if true manager field holds an AD object cn=name,ou=blah,dc=blah else it is a samaccount field
            // NULL not handeled as blanks
            // additionalFields takes the field names " table.field,"
            // Assumes table1 holds the correct data and returns a data reader with the update fields columns from table1
            // compare fields 1 & 2 should have the same number of items or it is likely that all rows will be found needing updating
            // returns fields from comparefields
            // returns the rows which table2 differs from table1

            // Assumes table1 holds the correct data and returns a data reader with the update fields columns from table1
            // returns the rows which table2's concatenated update fields differ from table1's concatenated update fields
            // eliminates rows which do not have a matching key in both tables
            //
            // additional fields table2.data2, table1.data5
            //*************************************************************************************************
            //| Table1                                  | Table2                                | Returned result
            //*************************************************************************************************
            //| ID            Data          Data5       | ID             Data       Data2       |                   | Table1.ID     Table1.DATA table1.data5        table2.data2
            //| 1             a             ty          | 1              a          e           | NOT RETURNED      |
            //| 2             b             e           | null           null       r           | NOT RETURNED      |              
            //| 3             c             uyt         | 3              null       f           | RETURNED          | 3             c           uyt                 f        
            //| 4             d             tr          | 4              e          w           | RETURNED          | 4             e           tr                  w
            //| 4             f             sr          | 4              f          w           | NOT RETURNED      |


            string compare1 = "";
            string compare2 = "";
            string fields = "";
            string notnull = "";
            string additionalfields = "";
            bool managerADtype = false;
            int i = 0;


            //int i;
            //string debugRecordCount = "";
            //string debug = "";                                     
            //debug = " total users from AD \n";
            //SqlCommand sqlDebugComm = new SqlCommand("select top 20 * FROM " + table2, sqlConn);
            //SqlDataReader debugReader = sqlDebugComm.ExecuteReader();
            //int debugFieldCount = debugReader.FieldCount;
            //for (i = 0; i < debugFieldCount; i++)
            //{
            //    debug += debugReader.GetName(i) + ", ";
            //}
            //debug += "\n";
            //while (debugReader.Read())
            //{
            //    for (i = 0; i < debugFieldCount; i++)
            //    {
            //        debug += (string)debugReader[i] + ", ";
            //    }
            //    debug += "\n";
            //}
            //sqlDebugComm = new SqlCommand("select count(" + pkey2 + ") FROM " + table1, sqlConn);
            //debugReader.Close();
            //debugRecordCount = sqlDebugComm.ExecuteScalar().ToString();
            //MessageBox.Show("table " + table2 + " has " + debugRecordCount + " records \n " + debugFieldCount + " fields \n sample data" + debug);
            
            //check to see what format the manager is in and what table it is in
            if (adField == 1)
            {
                foreach (string key in compareFields1)
                {
                    if (key == "CN") 
                    {
                    }
                    if (key == "manager")
                    {

                        SqlCommand sqlCheck = new SqlCommand("select top 1 manager FROM " + table1, sqlConn);
                        // create the command object
                        SqlDataReader checkReader;
                        try
                        {
                            sqlCheck.CommandTimeout = 360;
                            checkReader = sqlCheck.ExecuteReader();
                            log.addTrn(sqlCheck.CommandText.ToString(), "Query");
                            checkReader.Read();
                            if ((string)checkReader[0].ToString().Substring(0, 2) == "CN=")
                            {
                                managerADtype = true;
                            }
                            checkReader.Close();
                        }
                        catch (Exception ex)
                        {
                            log.addTrn("Failed SQL command " + sqlCheck.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                        }
                        
                    }
                }
            }
            else if (adField == 2)
            {
                foreach (string key in compareFields2)
                {
                    if (key == "manager")
                    {

                        SqlCommand sqlCheck = new SqlCommand("select top 1 manager FROM " + table2, sqlConn);
                        // create the command object
                        SqlDataReader checkReader;
                        try
                        {
                            sqlCheck.CommandTimeout = 360;
                            checkReader = sqlCheck.ExecuteReader();
                            log.addTrn(sqlCheck.ToString(), "Query");
                            checkReader.Read();
                            if ((string)checkReader[0].ToString().Substring(0, 2) == "CN=")
                            {
                                managerADtype = true;
                            }
                            checkReader.Close();
                        }
                        catch (Exception ex)
                        {
                            log.addTrn("Failed SQL command " + sqlCheck.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                        }
                    }
                }
            }

            // need a comand builder and research on the best way to compare all fields in a row
            // this basically will just issue a concatenation sql query to the DB for each field to compare
            //foreach (string key in compareFields1)
            //{
            //    if (managerADtype == false && key == "manager" && adField == 2)
            //    {
            //        //add code for substring of manager field
            //        compare1 = compare1 + "case when " + table2 + "." + compareFields2[i] + " <> '' then (substring(" + table1 + "." + key + ",4, charindex('ou=', " + table1 + "." + key + ")-5) COLLATE SQL_Latin1_General_CP1_CS_AS ) else '' end + ";
            //        fields += "substring(" + table1 + "." + key + ",4, charindex('ou=', " + table1 + "." + key + ")-5), ";
            //    }
            //    else
            //    {
            //         compare1 = compare1 + "case when " + table2 + "." + compareFields2[i] + " <> '' then (" + table1 + "." + key + " COLLATE SQL_Latin1_General_CP1_CS_AS ) else '' end + ";
            //         fields += table1 + "." + key + ", ";
            //    }
            //    i++;
            //}
            //i = 0;

            foreach (string key in compareFields1)
            {
                if (managerADtype == false && key == "manager" && adField == 2)
                {
                    //add code for substring of manager field
                    compare1 = compare1 + "substring(" + table1 + "." + key + ",4, charindex('ou=', " + table1 + "." + key + ")-5) COLLATE SQL_Latin1_General_CP1_CS_AS + ";
                    fields += "substring(" + table1 + "." + key + ",4, charindex('ou=', " + table1 + "." + key + ")-5), ";
                }
                else
                {
                    compare1 = compare1 + "ltrim(rtrim(" + table1 + "." + key + ")) COLLATE SQL_Latin1_General_CP1_CS_AS + ";
                    fields += table1 + "." + key + ", ";
                }
                i++;
            }
            i = 0;
            //compare1 = compare1 + table1 + "." + pkey1;
            foreach (string key in compareFields2)
            {
                if (managerADtype == false && key == "manager" && adField == 1)
                {
                    //add code for substring of manager field
                    compare2 = compare2 + "case when ltrim(rtrim(" + table2 + "." + compareFields2[i] + ")) <> '' then (substring(" + table2 + "." + key + ",4, charindex('ou=', " + table2 + "." + key + ")-5) COLLATE SQL_Latin1_General_CP1_CS_AS ) else '' end + ";
                    notnull += "case when len(ltrim(rtrim(" + table2 + "." + key + "))) > 3 then substring(" + table2 + "." + key + ",4, charindex('ou=', " + table2 + "." + key + ")-5) else '' end <> '' OR ";
                }
                else
                {
                    compare2 = compare2 + "case when ltrim(rtrim(" + table2 + "." + compareFields2[i] + ")) <> '' then (ltrim(rtrim(" + table2 + "." + key + ")) COLLATE SQL_Latin1_General_CP1_CS_AS ) else '' end + ";
                    //fields += table2 + "." + key + ", ";
                    notnull += "ltrim(rtrim(" + table2 + "." + key + ")) <> '' OR ";
                }
                i++;
            }
            //compare2 = compare2 + table2 + "." + pkey2;
            foreach (string key in additionalFields)
            {
                additionalfields += key;
            }
            // remove trailing comma and + 
            compare2 = compare2.Remove(compare2.Length - 2);
            compare1 = compare1.Remove(compare1.Length - 2);
            fields = fields.Remove(fields.Length - 2);
            notnull = notnull.Remove(notnull.Length - 3);
            additionalfields = additionalfields.Remove(additionalfields.Length - 2);
            SqlCommand sqlComm;
            if (additionalFields.Count > 0)
            {
                sqlComm = new SqlCommand("SELECT DISTINCT " /*+ compare2 + "," + compare1 + "," + table1 + "." + pkey1 + "," + table2 + "." + pkey2 + ","*/ + fields + ", " + additionalfields + " FROM " + table1 + " LEFT JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2 + " AND (" + compare2 + ") <> (" + compare1 + ") INNER JOIN " + table2 + " as [" + table2 + "temp] ON " + table1 + "." + pkey1 + " = [" + table2 + "temp]." + pkey2 + " AND " + table2 + "." + pkey2 + " IS NOT NULL WHERE " + notnull, sqlConn);
            }
            else
            {
                sqlComm = new SqlCommand("SELECT DISTINCT " + fields + " FROM " + table1 + " LEFT JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2 + " AND (" + compare2 + ") <> (" + compare1 + ") INNER JOIN " + table2 + " as [" + table2 + "temp] ON " + table1 + "." + pkey1 + " = [" + table2 + "temp]." + pkey2 + " AND " + table2 + "." + pkey2 + " IS NOT NULL WHERE " + notnull, sqlConn);
            }
            //AND " + table2 + "." + pkey2 + " != NULL
            try
            {
                sqlComm.CommandTimeout = 360;
                SqlDataReader r = sqlComm.ExecuteReader();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
                return r;
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
            return null;
        }
        public void CheckUpdateIntoNewTable(string table1, string table2, string pkey1, string pkey2, string newTable, ArrayList compareFields1, ArrayList compareFields2, SqlConnection sqlConn, LogFile log)
        {
            // Assumes table1 holds the correct data and returns a data reader with the update fields columns from table1
            // returns the rows which table2's concatenated update fields differ from table1's concatenated update fields
            // eliminates rows which do not have a matching key in both tables
            //*************************************************************************************************
            //| Table1                      | Table2                    | Returned result
            //*************************************************************************************************
            //| ID            Data          | ID             Data       |                   | Table1.ID     Table1.DATA
            //| 1             a             | 1              a          | NOT RETURNED      |
            //| 2             b             | null           null       | NOT RETURNED      |              
            //| 3             c             | 3              null       | RETURNED          | 3             c
            //| 4             d             | 4              e          | RETURNED          | 4             e
            //| 4             f             | 4              f          | NOT RETURNED      |

            string compare1 = "";
            string compare2 = "";
            string fields = "";
            // need a comand builder and research on the best way to compare all fields in a row
            // this basically will just issue a concatenation sql query to the DB for each field to compare
            foreach (string key in compareFields1)
            {
                if (key == pkey1)
                {
                    compare1 = compare1 + "ltrim(rtrim(" + table1 + "." + key + "))";
                }
                else
                {
                    compare1 = compare1 + "ltrim(rtrim(" + table1 + "." + key + ")) COLLATE SQL_Latin1_General_CP1_CS_AS + ";
                }
                fields += "ltrim(rtrim(" + table1 + "." + key + ")), ";
            }
            foreach (string key in compareFields2)
            {
                if (key == pkey1)
                {
                    compare2 = compare2 + "ltrim(rtrim(" + table2 + "." + key + "))";
                }
                else
                {
                    compare2 = compare2 + "ltrim(rtrim(" + table2 + "." + key + ")) COLLATE SQL_Latin1_General_CP1_CS_AS + ";
                }
                
            }
            // remove trailing comma and + 
            compare2 = compare2.Remove(compare2.Length - 2);
            compare1 = compare1.Remove(compare1.Length - 2);
            fields = fields.Remove(fields.Length - 2);
            SqlCommand sqlComm = new SqlCommand("SELECT DISTINCT " + fields + " INTO " + newTable + " FROM " + table1 + " LEFT JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2 + " AND (" + compare2 + ") = (" + compare1 + ") INNER JOIN " + table2 + " as [" + table2 + "temp] ON " + table1 + "." + pkey1 + " = [" + table2 + "temp]." + pkey2 + " WHERE " + table2 + "." + pkey2 + " IS NOT NULL", sqlConn);
            //AND " + table2 + "." + pkey2 + " != NULL
            try
            {
                sqlComm.CommandTimeout = 360;
                sqlComm.ExecuteNonQuery();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
        }
        public void SelectNicknamesClosestToActualNameIntoNewTable(string table1, string table2, string pkey1, string pkey2, string newTable, ArrayList selectFields, string targetField, ArrayList valueFields, SqlConnection sqlConn, LogFile log)
        {
            //select distinct soc_sec, 
            //(Select top 1 nickname
            //    from FHC_TEST_gmailNicknamesTable left join FHC_TEST_sqlusersTable on FHC_TEST_gmailNicknamesTable.soc_sec = FHC_TEST_sqlusersTable.sAMAccountName
            //    where soc_sec = a.soc_sec 
            //    order by soc_sec, 
            //        CHARINDEX(FHC_TEST_sqlusersTable.sn, FHC_TEST_gmailNicknamesTable.nickname) DESC,
            //        CHARINDEX(FHC_TEST_sqlusersTable.givenName, FHC_TEST_gmailNicknamesTable.nickname) DESC,
            //        CHARINDEX(FHC_TEST_sqlusersTable.middleName, FHC_TEST_gmailNicknamesTable.nickname) DESC
            //)
            //as nickname,
            //    (Select top 1 email
            //    from FHC_TEST_gmailNicknamesTable left join FHC_TEST_sqlusersTable on FHC_TEST_gmailNicknamesTable.soc_sec = FHC_TEST_sqlusersTable.sAMAccountName
            //    where soc_sec = a.soc_sec 
            //    order by soc_sec, CHARINDEX(FHC_TEST_sqlusersTable.sn, FHC_TEST_gmailNicknamesTable.nickname) DESC, CHARINDEX(FHC_TEST_sqlusersTable.givenName, FHC_TEST_gmailNicknamesTable.nickname) DESC, CHARINDEX(FHC_TEST_sqlusersTable.middleName, FHC_TEST_gmailNicknamesTable.nickname) DESC) as email
            //FROM FHC_TEST_gmailNicknamesTable as a
            //ORDER BY soc_sec;

            // Assumes table1 holds the correct data and returns a data reader with the update fields columns from table1
            // returns the rows which table2's concatenated update fields differ from table1's concatenated update fields
            // eliminates rows which do not have a matching key in both tables
            // adds convluted logic to deal with duplicates and select the one closest to the matching data from the table 2 assumed to be the correct first middle last name etc
            //*************************************************************************************************
            //| Table1                      | Table2                    | Returned result
            //*************************************************************************************************
            //| ID            Data          | ID             Data       |                   | Table1.ID     Table1.DATA
            //| 1             a             | 1              a          | NOT RETURNED      |
            //| 2             b             | null           null       | NOT RETURNED      |              
            //| 3             c             | 3              null       | RETURNED          | 3             c
            //| 4             d             | 4              e          | RETURNED          | 4             e

            //string table1 the table with the new nicknames
            //string table2 the table with the original names 
            //string pkey1
            //string pkey2
            //string newTable = ; the name of the new Table We want To return
            //ArrayList selectFields = new ArrayList( [ email] )List of fields we want returned fields should Be in table1
            //string targetField = "nickname"; name of the field we want to be close to ie like a google search result validity must be in table1
            //ArrayList valueFields = new ArrayList( [sn, givenName, middleName ] ); lift of fields which are being compared to the targetField to see how close it gets    


            string complexField = "";

            foreach (string key in selectFields)
            {
                complexField += "(SELECT TOP 1 " + table1 + "." + key;
                complexField += " FROM " + table1 + " LEFT JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2;
                complexField += " WHERE " + table1 + "." + pkey1 + " = extTable." + pkey1;
                complexField += " ORDER BY " + table2 + "." + pkey2 + ",";
                foreach (string value in valueFields)
                {
                    // should have first middle last as the arraylist for valueFields1
                    complexField += " CHARINDEX(" + table2 + "." + value + ", " + table1 + "." + targetField + ") DESC,";
                }
                // comma remove
                complexField = complexField.Remove(complexField.Length - 1);

                complexField += ") AS " + key + ",";
            }
            // comma remove
            complexField = complexField.Remove(complexField.Length - 1);

            SqlCommand sqlComm = new SqlCommand("SELECT DISTINCT extTable." + pkey1 + ", " + complexField + " INTO " + newTable + " FROM " + table1 + " AS extTable ORDER BY extTable." + pkey1, sqlConn);

            try
            {
                sqlComm.CommandTimeout = 360;
                sqlComm.ExecuteNonQuery();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
        }
        public void Levenshtein(string table1, string table2, string pkey1, string pkey2, string newTable, ArrayList selectFields, string targetField, ArrayList valueFields, SqlConnection sqlConn, LogFile log)
        {
    
            //SELECT dbo.LEVENSHTEIN((FHC_LDAP_sqlusersTable.givenName + '.' + FHC_LDAP_sqlusersTable.sn), FHC_LDAP_gmailNicknamesTable.nickname),
            //(FHC_LDAP_sqlusersTable.givenName + '.' + FHC_LDAP_sqlusersTable.sn) as templateName,
            //FHC_LDAP_gmailNicknamesTable.nickname,
            //FHC_LDAP_gmailNicknamesTable.Email, 
            //FHC_LDAP_gmailNicknamesTable.soc_sec 
            //FROM FHC_LDAP_gmailNicknamesTable LEFT JOIN FHC_LDAP_sqlusersTable ON FHC_LDAP_gmailNicknamesTable.soc_sec = FHC_LDAP_sqlusersTable.sAMAccountName 
            //group by FHC_LDAP_gmailNicknamesTable.soc_sec, FHC_LDAP_gmailNicknamesTable.Email, FHC_LDAP_gmailNicknamesTable.nickname, 
            //(FHC_LDAP_sqlusersTable.givenName + '.' + FHC_LDAP_sqlusersTable.sn)
            //order by FHC_LDAP_gmailNicknamesTable.soc_sec

            //uses levensthein to calculate the closest nickname to the first.last and returns that nickname

            //string table1 the table with the new nicknames
            //string table2 the table with the original names 
            //string pkey1
            //string pkey2
            //string newTable = ; the name of the new Table We want To return
            //ArrayList selectFields = new ArrayList( [ email] )List of fields we want returned fields should Be in table1
            //string targetField = "nickname"; name of the field we want to be close to ie like a google search result validity must be in table1
            //ArrayList valueFields = new ArrayList( [sn, givenName, middleName ] ); lift of fields which are being compared to the targetField to see how close it gets    


            string complexField = "";

            foreach (string key in selectFields)
            {
                complexField += "(SELECT TOP 1 " + table1 + "." + key;
                complexField += " FROM " + table1 + " LEFT JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2;
                complexField += " WHERE " + table1 + "." + pkey1 + " = extTable." + pkey1;
                complexField += " ORDER BY " + table2 + "." + pkey2 + ",";
                foreach (string value in valueFields)
                {
                    // should have first middle last as the arraylist for valueFields1
                    complexField += " CHARINDEX(" + table2 + "." + value + ", " + table1 + "." + targetField + ") DESC,";
                }
                // comma remove
                complexField = complexField.Remove(complexField.Length - 1);

                complexField += ") AS " + key + ",";
            }
            // comma remove
            complexField = complexField.Remove(complexField.Length - 1);

            SqlCommand sqlComm = new SqlCommand("SELECT DISTINCT extTable." + pkey1 + ", " + complexField + " INTO " + newTable + " FROM " + table1 + " AS extTable ORDER BY extTable." + pkey1, sqlConn);

            try
            {
                sqlComm.CommandTimeout = 360;
                sqlComm.ExecuteNonQuery();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
        }

        public void CheckEmailUpdateIntoNewTable(string email1, string email2, string table1, string table2, string pkey1, string pkey2, string newTable, ArrayList compareFields1, ArrayList compareFields2, SqlConnection sqlConn, LogFile log)
        {

            //select FHC_test2_gmailNicknamesTable.soc_sec, FHC_test2_gmailNicknamesTable.Email
            //FROM FHC_test2_gmailNicknamesTable INNER JOIN FHC_test2_sqlNicknamesTable 
            //ON FHC_test2_gmailNicknamesTable.soc_sec = FHC_test2_sqlNicknamesTable.soc_sec 
            //where FHC_test2_gmailNicknamesTable.Email not in (select FHC_test2_gmailnicknamestable.email from FHC_test2_gmailnicknamestable)

            // Assumes table1 holds the correct data and returns a data reader with the update fields columns from table1
            // returns the rows which table2's concatenated update fields differ from table1's concatenated update fields
            // eliminates rows which do not have a matching key in both tables
            //*************************************************************************************************
            //| Table1                      | Table2                    | Returned result
            //*************************************************************************************************
            //| ID            Data          | ID             Data       |                   | Table1.ID     Table1.DATA
            //| 1             a             | 1              a          | NOT RETURNED      |
            //| 2             b             | null           null       | NOT RETURNED      |              
            //| 3             c             | 3              null       | RETURNED          | 3             c
            //| 4             d             | 4              e          | RETURNED          | 4             e

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
            SqlCommand sqlComm = new SqlCommand("SELECT DISTINCT " + fields + " INTO " + newTable + " FROM " + table1 + " INNER JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2 + " WHERE " + table1 + "." + email1 + " NOT IN ( SELECT " + table2 + "." + email2 + " FROM " + table2 + " )", sqlConn);
            //AND " + table2 + "." + pkey2 + " != NULL
            try
            {
                sqlComm.CommandTimeout = 360;
                sqlComm.ExecuteNonQuery();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
        }


        public ArrayList SqlColumns(UserSynch userconfig)
        {
            ArrayList columnList = new ArrayList();
            if (userconfig.DBCatalog != "" && userconfig.DataServer != "")
            {
                //populates columns dialog with columns depending on the results of a query
                try
                {
                    SqlConnection sqlConn = new SqlConnection("Data Source=" + userconfig.DataServer.ToString() + ";Initial Catalog=" + userconfig.DBCatalog.ToString() + ";Integrated Security=SSPI;");

                    sqlConn.Open();
                    // create the command object
                    SqlCommand sqlComm = new SqlCommand("SELECT column_name FROM information_schema.columns WHERE table_name = '" + userconfig.User_dbTable + "'", sqlConn);
                    sqlComm.CommandTimeout = 360;
                    SqlDataReader r = sqlComm.ExecuteReader();
                    while (r.Read())
                    {
                        columnList.Add((string)r[0].ToString().Trim());
                    }
                    r.Close();
                    sqlConn.Close();
                }
                catch
                {
                    columnList.Add("Error pulling SQL columns");
                }
            }
            return columnList;
        }

        // SQL update tools
        public void Mass_update_email_field(DataTable users, SqlConnection sqlConn, GmailUsers gusersyn, LogFile log)
        {
            // string concatenation replaced with stringbuilder due to rumored performance increases
            // 
            /*
           UPDATE address
           SET e_mail2 = e_mail,
               e_mail = 'actual address'
           WHERE primary_key = '####' AND where clause
              
             * customized to work with GmailUsers object pulling the email field and primary keys for the transition form there
              
            */

            int i;
            StringBuilder sqlstring = new StringBuilder();
            SqlCommand sqlComm;



            for (i = 0; i < users.Rows.Count; i++)
            {
                // now add the data
                sqlstring.Remove(0, sqlstring.Length);
                sqlstring.Append("UPDATE " + gusersyn.Writeback_table + " ");
                if (gusersyn.Writeback_transfer_email_checkbox == true)
                {
                    sqlstring.Append("SET " + gusersyn.Writeback_secondary_email_field + " = " + gusersyn.Writeback_email_field + ",");
                    sqlstring.Append(gusersyn.Writeback_email_field + " = '" + users.Rows[i][gusersyn.Writeback_email_field].ToString().Replace("'", "''") + "' ");
                }
                else
                {
                    sqlstring.Append("SET " + gusersyn.Writeback_email_field + " = '" + users.Rows[i][gusersyn.Writeback_email_field].ToString().Replace("'", "''") + "' ");
                }

                sqlstring.Append("WHERE " + gusersyn.Writeback_primary_key + " = '" + users.Rows[i][gusersyn.Writeback_primary_key].ToString().Replace("'", "''") + "' ");
                if (gusersyn.Writeback_where_clause.ToString().Trim() != "")
                {
                    sqlstring.Append("AND " + gusersyn.Writeback_where_clause);
                }
                sqlComm = new SqlCommand(sqlstring.ToString(), sqlConn);


                // MessageBox.Show(sqlstring.Length.ToString());
                try
                {
                    sqlComm.CommandTimeout = 360;
                    sqlComm.ExecuteNonQuery();
                    log.addTrn(sqlComm.CommandText.ToString(), "Query");
                    log.addTrn("DB email writeback, user " + users.Rows[i][gusersyn.Writeback_primary_key].ToString().Replace("'", "''") + ", email " + users.Rows[i][gusersyn.Writeback_email_field].ToString().Replace("'", "''"), "Transaction");
                }
                catch (Exception ex)
                {
                    log.addTrn("DB email writeback failure, user " + users.Rows[i][gusersyn.Writeback_primary_key].ToString().Replace("'", "''") + ", email " + users.Rows[i][gusersyn.Writeback_email_field].ToString().Replace("'", "''") + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                }
            }


        }
        public void Mass_update_email_field(SqlDataReader users, SqlConnection sqlConn, GmailUsers gusersyn, LogFile log)
        {
            // string concatenation replaced with stringbuilder due to rumored performance increases
            // 
            /*
           UPDATE address
           SET e_mail2 = e_mail,
               e_mail = 'actual address'
           WHERE primary_key = '####' AND where clause
              
             * customized to work with GmailUsers object pulling the email field and primary keys for the transition form there
              
            */

            StringBuilder sqlstring = new StringBuilder();
            SqlCommand sqlComm;


            try
            {
                while (users.Read())
                {
                    // now add the data
                    sqlstring.Remove(0, sqlstring.Length);
                    sqlstring.Append("UPDATE " + gusersyn.Writeback_table + " ");
                    if (gusersyn.Writeback_transfer_email_checkbox == true)
                    {
                        sqlstring.Append("SET " + gusersyn.Writeback_secondary_email_field + " = " + gusersyn.Writeback_email_field + ",");
                        sqlstring.Append(gusersyn.Writeback_email_field + " = '" + users[gusersyn.Writeback_email_field].ToString().Replace("'", "''") + "' ");
                    }
                    else
                    {
                        sqlstring.Append("SET " + gusersyn.Writeback_email_field + " = '" + users[gusersyn.Writeback_email_field].ToString().Replace("'", "''") + "' ");
                    }

                    sqlstring.Append("WHERE " + gusersyn.Writeback_primary_key + " = '" + users[gusersyn.Writeback_primary_key].ToString().Replace("'", "''") + "' ");
                    if (gusersyn.Writeback_where_clause.ToString().Trim() != "")
                    {
                        sqlstring.Append("AND " + gusersyn.Writeback_where_clause);
                    }
                    sqlComm = new SqlCommand(sqlstring.ToString(), sqlConn);


                    // MessageBox.Show(sqlstring.Length.ToString());
                    try
                    {
                        sqlComm.CommandTimeout = 360;
                        sqlComm.ExecuteNonQuery();
                        log.addTrn(sqlComm.CommandText.ToString(), "Query");
                        log.addTrn("DB email writeback, user " + users[gusersyn.Writeback_primary_key].ToString().Replace("'", "''") + ", email " + users[gusersyn.Writeback_email_field].ToString().Replace("'", "''"), "Transaction");
                    }
                    catch (Exception ex)
                    {
                        log.addTrn("DB email writeback failure, user " + users[gusersyn.Writeback_primary_key].ToString().Replace("'", "''") + ", email " + users[gusersyn.Writeback_email_field].ToString().Replace("'", "''") + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                    }
                    sqlComm.Dispose();
                }
            }
            catch (Exception ex)
            {
                log.addTrn("Issue in DB writeback datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }


        }
        public void Mass_Table_update(string sourceTable, string targetTable, string sourcePimaryKey, string targetPrimaryKey, ArrayList sourceColumns, ArrayList targetColumns, string whereClause, SqlConnection sqlConn, LogFile log)
        {
            // sourceColumns and targetColumns must have the same number of columns (its a one to one transfer)
            // updates the targetTable's targetColumns with the data from the sourceTable's  sourceColumns
            //
            //UPDATE table1
            //        SET table1.col = table2.col1
            //FROM table1 INNER JOIN table2 ON table1.Pkey = table2.Pkey
            //WHERE variable clause

            //UPDATE address 
            //SET address.gmail = name.gender 
            //FROM address INNER JOIN name ON address.soc_sec = name.soc_sec
            //WHERE address.preferred = 1

            string columnValues = "";
            int sourceCount = sourceColumns.Count;
            int i = 0;
            for (i = 0; i < sourceCount; i++)
            {
                columnValues += targetTable + "." + targetColumns[i].ToString() + " = " + sourceTable + "." + sourceColumns[i].ToString() + ", ";
            }

            columnValues = columnValues.Remove(columnValues.Length - 2);
            SqlCommand sqlComm;
            if (whereClause.Length == 0)
            {
                sqlComm = new SqlCommand("UPDATE " + targetTable + " SET " + columnValues + " FROM " + targetTable + " INNER JOIN " + sourceTable + " ON " + targetTable + "." + targetPrimaryKey + " = " + sourceTable + "." + sourcePimaryKey, sqlConn);
            }
            else
            {
                sqlComm = new SqlCommand("UPDATE " + targetTable + " SET " + columnValues + " FROM " + targetTable + " INNER JOIN " + sourceTable + " ON " + targetTable + "." + targetPrimaryKey + " = " + sourceTable + "." + sourcePimaryKey + " WHERE " + whereClause, sqlConn);
            }
            try
            {
                sqlComm.CommandTimeout = 360;
                sqlComm.ExecuteNonQuery();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
        }
        public void Mass_Email_Shift(GmailUsers gusersyn, string sourceTable, string targetTable, string sourcePimaryKey, string targetPrimaryKey, ArrayList sourceColumnsAndTable, ArrayList targetColumnsAndTable, string whereClause, SqlConnection sqlConn, LogFile log)
        {
            // sourceColumnsAndTable and targetColumnsAndTable must have the same number of columns (its a one to one transfer)
            // updates the targetColumnsAndTable with the data from the sourceColumnsAndTable
            // does not overwrite secondary fields which contain the email domain from the gusersyn
            //
            // sourceColumnsAndTable expects   "table.column"
            // targetColumnsAndTable expects   "table.column"

            // massUpdateSameTable  truely a useless flag to overload the function
            //
            //UPDATE table1
            //        SET table1.col = table2.col1
            //FROM table1 INNER JOIN table2 ON table1.Pkey = table2.Pkey
            //WHERE variable clause

            //UPDATE a1 
            //SET a1.e_mail2 = a1.e_mail 
            //FROM address as a1 INNER JOIN (select * from address as a2 where rtrim(a2.gmail) <> '') as a2 ON a1.soc_sec = a2.soc_sec 


            string columnValues = "";
            int sourceCount = sourceColumnsAndTable.Count;
            int i = 0;
            for (i = 0; i < sourceCount; i++)
            {
                columnValues += targetColumnsAndTable[i].ToString() + " = " + sourceColumnsAndTable[i].ToString() + ", ";
            }

            columnValues = columnValues.Remove(columnValues.Length - 2);
            SqlCommand sqlComm;
            if (whereClause.Length == 0)
            {
                sqlComm = new SqlCommand("UPDATE " + targetTable + " SET " + columnValues + " FROM " + targetTable + " INNER JOIN " + sourceTable + " ON " + targetTable + "." + targetPrimaryKey + " = " + sourceTable + "." + sourcePimaryKey + " WHERE " + sourceColumnsAndTable[0] + " not like '%" + gusersyn.Admin_domain + "%' AND " + sourceColumnsAndTable[0] + " <> '' AND " + sourceColumnsAndTable[0] + " <> '?' AND " + sourceColumnsAndTable[0] + " IS NOT NULL ", sqlConn);
            }
            else
            {
                sqlComm = new SqlCommand("UPDATE " + targetTable + " SET " + columnValues + " FROM " + targetTable + " INNER JOIN " + sourceTable + " ON " + targetTable + "." + targetPrimaryKey + " = " + sourceTable + "." + sourcePimaryKey + " WHERE " + sourceColumnsAndTable[0] + " not like '%" + gusersyn.Admin_domain + "%' AND " + sourceColumnsAndTable[0] + " <> '' AND " + sourceColumnsAndTable[0] + " <> '?' AND " + sourceColumnsAndTable[0] + " IS NOT NULL AND " + whereClause, sqlConn);
            }
            try
            {
                sqlComm.CommandTimeout = 360;
                sqlComm.ExecuteNonQuery();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
        }



        // Gmail tools
        public string GetNewUserNickname(AppsService service, string studentID, string firstName, string midName, string lastName, int i, bool complete)
        {
            // this could really get screwed if there are enough duplicates it will be only do first.m.last f.middle.last
            // i can be used to start the process somewhere in the middle of the name
            string returnvalue = "";
            firstName = firstName.Replace(".", "");
            lastName = lastName.Replace(".", "");
            midName = midName.Replace(".", "");
            while (complete == false)
            {
                int r = midName.Length + 1;

                if (i == 0)
                {
                    returnvalue = firstName + "." + lastName;
                }
                else
                {
                    if (i < r)
                    {
                        returnvalue = firstName + "." + midName.Substring(0, i) + "." + lastName;
                    }
                    else
                    {
                        if (i < (r + firstName.Length))
                        {
                            returnvalue = firstName.Substring(0, (i - r)) + "." + midName + "." + lastName;
                        }
                    }
                }
                if (i > (r + firstName.Length))
                {
                    returnvalue = "failure";
                    complete = true;
                }


                try
                {
                    if (complete == false)
                    {
                        returnvalue = Regex.Replace(Regex.Replace(returnvalue, @"[^a-z|^A-Z|^0-9|^\.|_|-]|[\^|\|]", ""), @"\.+", ".");
                        service.CreateNickname(studentID, returnvalue);
                        complete = true;
                    }
                }
                catch (AppsException apex)
                {
                    //MessageBox.Show("Nickname apps exception " + apex.ErrorCode.ToString() + "  +++  \n" + apex.Data.ToString() + "  +++  \n" + apex.Message.ToString() + "  +++  \n" + apex.Reason.ToString() + "  +++  \n" + apex.Source.ToString());
                    if (apex.ErrorCode == "1301")
                    {
                        // this error about a non existent entry seems to indicate the nickname is already created
                        complete = true;
                    }
                    i++;
                }
                catch (Exception)
                {
                    //MessageBox.Show("Nickname issue " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
                    i++;
                    complete = false;
                }
            }
            return returnvalue.Trim();
        }
        public DataTable Create_Gmail_Users(AppsService service, GmailUsers gusersyn, SqlDataReader users, LogFile log)
        {
            // user alising not created yet
            // Takes the SQLDataReader and creates all users in the reader

            // create the table for holding the users final gmail account infomation for email writeback
            DataTable returnvalue = new DataTable();
            DataRow row;

            returnvalue.TableName = "users";

            returnvalue.Columns.Add(gusersyn.Writeback_primary_key);
            returnvalue.Columns.Add(gusersyn.Writeback_email_field);


            row = returnvalue.NewRow();
            string studentID = "";
            string first_name = "";
            string last_name = "";
            string middle_name = "";
            string password = "";
            string userNickName = "Aliasing off";
            try
            {
                while (users.Read())
                {

                    try
                    {
                        // using _ as escape character allows illegal characters in username
                        studentID = System.Web.HttpUtility.UrlEncode(users[gusersyn.User_StuID].ToString()).Replace("+", " ").Replace("*", "%2A").Replace("!", "%21").Replace("(", "%28").Replace(")", "%29").Replace("'", "%27").Replace("_", "%5f").Replace(" ", "%20").Replace("%", "_");
                        // names are less restricitve the only illegal characters are < > =
                        first_name = users[gusersyn.User_Fname].ToString().Replace("<", "%3c").Replace(">", "%3e").Replace("=", "%3d").Replace("%", "%25");
                        middle_name = users[gusersyn.User_Mname].ToString().Replace("<", "%3c").Replace(">", "%3e").Replace("=", "%3d").Replace("%", "%25");
                        last_name = users[gusersyn.User_Lname].ToString().Replace("<", "%3c").Replace(">", "%3e").Replace("=", "%3d").Replace("%", "%25");
                        if (gusersyn.User_password_generate_checkbox == false)
                        {
                            // password needs to bea able to handle special characters
                            password = users[gusersyn.User_password].ToString();
                        }
                        else
                        {
                            password = GetPassword(14);
                        }
                        if (gusersyn.User_password_short_fix_checkbox == true && password.Length < 8)
                        {
                            password = GetPassword(14);
                        }



                        //Create a new user.
                        UserEntry insertedEntry = service.CreateUser(studentID, first_name, last_name, password);

                        //if (gusersyn.Levenshtein == true)
                        //{
                            // create user ailas here
                            userNickName = GetNewUserNickname(service, studentID, first_name, middle_name, last_name, 0, false);

                            row[0] = studentID;
                            if (userNickName != "failure")
                            {
                                row[1] = userNickName + "@" + gusersyn.Admin_domain;
                            }
                            else
                            {
                                row[1] = studentID + "@" + gusersyn.Admin_domain;
                            }

                            returnvalue.Rows.Add(row);
                            row = returnvalue.NewRow();

                            log.addTrn("Added Gmail user " + studentID + "@" + gusersyn.Admin_domain + " Aliased as " + userNickName + "@" + gusersyn.Admin_domain, "Transaction");
                       // }
                    }
                    catch (AppsException e)
                    {
                        log.addTrn("Failed adding Gmail user " + studentID + "@" + gusersyn.Admin_domain + " Aliased as " + userNickName + "@" + gusersyn.Admin_domain + " failed " + e.Message.ToString() + " reason " + e.Reason.ToString(), "Error");
                    }
                    catch (Exception ex)
                    {
                        log.addTrn("Failed adding Gmail user " + studentID + "@" + gusersyn.Admin_domain + " Aliased as " + userNickName + "@" + gusersyn.Admin_domain + " failed " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                    }
                }
            }
            catch (Exception ex)
            {
                log.addTrn("Issue adding gmail users datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
            return returnvalue;

        }
        public void DeleteGmailUserAccount(AppsService service, string userID, LogFile log)
        {
            try
            {
                service.DeleteUser(userID);
            }
            catch (Exception ex)
            {
                log.addTrn("Failed Delete gmail account " + userID + " exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }


        }
        public void UpdateGmailUser(AppsService service, GmailUsers gusersyn, SqlDataReader usersToUpdate, LogFile log)
        {
            string userNickName = "";
            string middlename = "";
            try
            {
                while (usersToUpdate.Read())
                {
                    try
                    {
                        UserEntry gmailUser = service.RetrieveUser((string)usersToUpdate[gusersyn.User_StuID]);
                        //special gmail username replace string only allows -_. special character thru
                        gmailUser.Login.UserName = System.Web.HttpUtility.UrlEncode(usersToUpdate[gusersyn.User_StuID].ToString()).Replace("+", " ").Replace("*", "%2A").Replace("!", "%21").Replace("(", "%28").Replace(")", "%29").Replace("'", "%27").Replace("_", "%5f").Replace(" ", "%20").Replace("%", "_");
                        gmailUser.Name.FamilyName = usersToUpdate[gusersyn.User_Lname].ToString().Replace("<", "%3c").Replace(">", "%3e").Replace("=", "%3d").Replace("%", "%25");
                        gmailUser.Name.GivenName = usersToUpdate[gusersyn.User_Fname].ToString().Replace("<", "%3c").Replace(">", "%3e").Replace("=", "%3d").Replace("%", "%25");
                        middlename = usersToUpdate[gusersyn.User_Mname].ToString().Replace("<", "%3c").Replace(">", "%3e").Replace("=", "%3d").Replace("%", "%25");
                        service.UpdateUser(gmailUser);
                        log.addTrn("Updated " + System.Web.HttpUtility.UrlEncode(usersToUpdate[gusersyn.User_StuID].ToString()).Replace("+", " ").Replace("*", "%2A") + " because of name change. New Name is " + gmailUser.Name.FamilyName.ToString() + ", " + gmailUser.Name.GivenName.ToString(), "Transaction");
                       // if (gusersyn.Levenshtein == true)
                       // {
                            userNickName = GetNewUserNickname(service, gmailUser.Login.UserName, gmailUser.Name.GivenName, middlename, gmailUser.Name.FamilyName, 0, false);
                            log.addTrn("Added New Alias for " + gmailUser.Login.UserName + "@" + gusersyn.Admin_domain + " Aliased as " + userNickName + "@" + gusersyn.Admin_domain, "Transaction");
                       // }
                        
                    }
                    catch (Exception ex)
                    {
                        log.addTrn("Failed update gmail account " + System.Web.HttpUtility.UrlEncode(usersToUpdate[gusersyn.User_StuID].ToString()).Replace("+", " ").Replace("*", "%2A").Replace("!", "%21").Replace("(", "%28").Replace(")", "%29").Replace("'", "%27").Replace("_", "%5f").Replace(" ", "%20").Replace("%", "_") + " exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                    }

                }
            }
            catch (Exception ex)
            {
                log.addTrn("Issue updating gmail users datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
        }
        public void CreateSendAs(GmailUsers gusersyn, SqlDataReader userNicknames, string sendASFieldName, string replyToFieldName, LogFile log)
        {
            //AppsService service = new AppsService(gusersyn.Admin_domain, gusersyn.Admin_user + "@" + gusersyn.Admin_domain, gusersyn.Admin_password);
            GoogleMailSettingsService gmailSettings = new GoogleMailSettingsService(gusersyn.Admin_domain, gusersyn.Admin_domain);
            gmailSettings.setUserCredentials(gusersyn.Admin_user + "@" + gusersyn.Admin_domain, gusersyn.Admin_password);
            try
            {
                while (userNicknames.Read())
                {
                    try
                    {
                        gmailSettings.CreateSendAs((string)userNicknames[gusersyn.User_StuID], (string)userNicknames[gusersyn.User_Fname] + " " + (string)userNicknames[gusersyn.User_Lname], (string)userNicknames[sendASFieldName], (string)userNicknames[replyToFieldName], "true");
                        log.addTrn("Created send as alias " + (string)userNicknames[sendASFieldName] + " for userlogin " + (string)userNicknames[gusersyn.User_StuID], "Transaction");
                    }
                    catch (Exception ex)
                    {
                        log.addTrn("Failed user send as creation " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                    }
                }
            }
            catch (Exception ex)
            {
                log.addTrn("Issue creating send as datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
        }

        // Gmail data pull
        public DataTable Get_Gmail_Users(AppsService service, GmailUsers gusersyn, string table, LogFile log)
        {
            // nicknames will have to dealt with seperately
            DataTable returnvalue = new DataTable();
            DataRow row;

            returnvalue.TableName = table;

            int i = 0;
            int count = 0;
            returnvalue.Columns.Add(gusersyn.User_StuID);
            returnvalue.Columns.Add(gusersyn.User_Fname);
            returnvalue.Columns.Add(gusersyn.User_Lname);


            try
            {
                UserFeed usersList = service.RetrieveAllUsers();
                count = usersList.Entries.Count;
                //result.AppendText("domain " + service.Domain + "\n");
                //result.AppendText("app name " + service.ApplicationName + "\n");
                //result.AppendText("users " + count + "\n");
                row = returnvalue.NewRow();
                for (i = 0; i < count; i++)
                {
                    UserEntry userEntry = usersList.Entries[i] as UserEntry;
                    // special handling for userID due to % being an illegal character using _ as an escape character
                    row[0] = (System.Web.HttpUtility.UrlDecode(userEntry.Login.UserName.ToString().Replace("_", "%")));
                    // decode names due to encoding to remove <>= characters
                    row[1] = (System.Web.HttpUtility.UrlDecode(userEntry.Name.GivenName.ToString()));
                    row[2] = (System.Web.HttpUtility.UrlDecode(userEntry.Name.FamilyName.ToString()));


                    //userList.Add(userEntry.Login.UserName.ToString());

                    returnvalue.Rows.Add(row);
                    row = returnvalue.NewRow();
                }
            }
            catch (Exception ex)
            {
                log.addTrn("failed to pull gmail user list exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
            return returnvalue;
        }
        public DataTable Get_Gmail_Nicknames(AppsService service, GmailUsers gusersyn, string table, LogFile log)
        {
            // nicknames retrieval
            DataTable returnvalue = new DataTable();
            DataRow row;

            returnvalue.TableName = table;

            int i = 0;
            int count = 0;
            string nickname = "";
            returnvalue.Columns.Add(gusersyn.Writeback_primary_key);
            returnvalue.Columns.Add("nickname");
            returnvalue.Columns.Add("Email");

            try
            {
                NicknameFeed usersNicknameList = service.RetrieveAllNicknames();
                count = usersNicknameList.Entries.Count;
                //result.AppendText("domain " + service.Domain + "\n");
                //result.AppendText("app name " + service.ApplicationName + "\n");
                //result.AppendText("users " + count + "\n");
                row = returnvalue.NewRow();
                for (i = 0; i < count; i++)
                {
                    NicknameEntry userNicknameEntry = usersNicknameList.Entries[i] as NicknameEntry;
                    // special handling for userID due to % being an illegal character using _ as an escape character
                    row[0] = (System.Web.HttpUtility.UrlDecode(userNicknameEntry.Login.UserName.ToString().Replace("_", "%")));
                    // decode names due to encoding to remove <>= characters
                    nickname = (System.Web.HttpUtility.UrlDecode(userNicknameEntry.Nickname.Name.ToString()));
                    row[1] = nickname;
                    row[2] = nickname + "@" + gusersyn.Admin_domain.ToString();

                    returnvalue.Rows.Add(row);
                    row = returnvalue.NewRow();
                }
            }
            catch (Exception ex)
            {
                log.addTrn("failed to pull gmail nickname list exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
            return returnvalue;
        }


        // Strong Password tool
        public string GetPassword(int length)
        {
            int i = 0;
            char[] charslow = new char[] { 'q', 'w', 'e', 'r', 't', 'y', 'u', 'i', 'o', 'p', 'a', 's', 'd', 'f', 'g', 'h', 'j', 'k', 'l', 'z', 'x', 'c', 'v', 'b', 'n', 'm' };
            char[] charsup = new char[] { 'Q', 'W', 'E', 'R', 'T', 'Y', 'U', 'I', 'O', 'P', 'A', 'S', 'D', 'F', 'G', 'H', 'J', 'K', 'L', 'Z', 'X', 'C', 'V', 'B', 'N', 'M' };
            char[] special = new char[] { '{', '}', '.', '!', '#', '$', '%', '^', '&', '(', ')', '-', '_', '~' };
            char[] number = new char[] { '1', '2', '3', '4', '5', '6', '7', '8', '9', '0' };
            int value = 0;
            int j = 0;
            int sumpicked = 0;
            int[] picked = new int[] { 0, 0, 0, 0 };
            string returnvalue = "";
            Random rndnum = new Random();
            for (i = 0; i < length; i++)
            {
                sumpicked = 0;
                for (j = 0; j < 4; j++)
                {
                    sumpicked += picked[j];
                }
                if (sumpicked < 3 && (length - (i + 3)) < 0)
                {
                    while (picked[value] == 1)
                    {
                        value = rndnum.Next(0, 4);
                    }
                }
                else
                {
                    value = rndnum.Next(0, 4);
                }
                switch (value)
                {
                    case 0:
                        returnvalue += charsup[rndnum.Next(0, 25)];
                        picked[0] = 1;
                        break;
                    case 1:
                        returnvalue += charslow[rndnum.Next(0, 25)];
                        picked[1] = 1;
                        break;
                    case 2:
                        returnvalue += special[rndnum.Next(0, 14)];
                        picked[2] = 1;
                        break;
                    case 3:
                        returnvalue += number[rndnum.Next(0, 9)];
                        picked[3] = 1;
                        break;
                }
            }
            return returnvalue;
        }


        // Log file utilities
        public void savelog(LogFile log, ConfigSettings settingsConfig)
        {
            // create a file stream, where "c:\\testing.txt" is the file path
            if (settingsConfig.LogType == "Text File")
            {
                string datetimeappend = DateTime.Today.Date.ToString() + DateTime.Today.TimeOfDay.ToString();
                System.IO.FileStream fs = new System.IO.FileStream(settingsConfig.LogDirectory + datetimeappend, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write, System.IO.FileShare.ReadWrite);

                // create a stream writer
                System.IO.StreamWriter sw = new System.IO.StreamWriter(fs, System.Text.Encoding.ASCII);

                StringBuilder result = new StringBuilder();

                foreach (DataColumn dc in log.logTrns.Columns)
                {
                    result.AppendFormat("{0}\t\t\t\t", dc.ColumnName);
                }

                result.Append("\r\n");

                foreach (DataRow dr in log.logTrns.Rows)
                {
                    foreach (DataColumn dc in log.logTrns.Columns)
                    {
                        result.AppendFormat("{0}\t\t\t\t", (Convert.IsDBNull(dr[dc.ColumnName]) ? string.Empty : dr[dc.ColumnName].ToString()));
                    }

                    result.Append("\r\n");
                }



                sw.Write(result.ToString());

                // flush buffer (so the text really goes into the file)
                sw.Flush();

                // close stream writer and file
                sw.Close();
                fs.Close();
            }
            if (settingsConfig.LogType == "Database")
            {
                //create sql for log file if it does not exist
                // sqlConn must be an open connection
                string table = "FHC_LOG_ldap_magic";
                DataTable data = new DataTable();
                log.initiateTrn();


                StringBuilder sqlstring = new StringBuilder();
                SqlConnection sqlConn = new SqlConnection("Data Source=" + settingsConfig.LogDB + ";Initial Catalog=" + settingsConfig.LogCatalog + ";Integrated Security=SSPI;Connect Timeout=360;");
                sqlConn.Open();
                SqlCommand sqlComm;

                sqlstring.Append("CREATE TABLE [" + table + "]([Message] [text], [Type] [varchar](50), [Timestamp] [datetime] NULL) ON [PRIMARY]");
                sqlComm = new SqlCommand(sqlstring.ToString(), sqlConn);
                try
                {
                    sqlComm.CommandTimeout = 360;
                    sqlComm.ExecuteNonQuery();
                    log.addTrn(sqlComm.CommandText.ToString(), "Query");
                    log.addTrn("table created " + table, "Transaction");
                }
                catch (Exception ex)
                {
                    log.addTrn("Table already exists or Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                    //log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
                }

                // copy data into table
                try
                {
                    SqlBulkCopy sbc = new SqlBulkCopy(sqlConn);
                    sbc.DestinationTableName = table;
                    sbc.WriteToServer(log.logTrns);
                    sbc.Close();
                }
                catch (Exception ex)
                {
                    log.addTrn("Failed SQL bulk copy " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                }
                //create datatable with log information
                //use append table to sql bulk copy to table
                //optionally create blank table and use copy into/merge to get records into existing table
            }

        }


        // possible not in use

        public void MoveADObject(string objectLocation, string newLocation, LogFile log)
        {
            //For brevity, removed existence checks
            // EXPECTS FULL Distinguished Name for both variables "LDAP://CN=xxx,DC=xxx,DC=xxx"
            try
            {
                DirectoryEntry eLocation = new DirectoryEntry(objectLocation);
                DirectoryEntry nLocation = new DirectoryEntry(newLocation);
                string newName = eLocation.Name;
                eLocation.MoveTo(nLocation, newName);
                nLocation.Close();
                eLocation.Close();
            }
            catch (Exception ex)
            {
                log.addTrn("failed to move ad object from " + objectLocation + " to " + newLocation + " exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }

        }
        public string GetAttributeValuesSingleString(string attributeName, string objectDn, LogFile log)
        {
            string strValue;
            try
            {
                DirectoryEntry ent = new DirectoryEntry(objectDn);
                strValue = ent.Properties[attributeName].Value.ToString();
                ent.Close();
                ent.Dispose();
            }
            catch
            {
                log.addTrn("failed to pull " + attributeName + " on object " + objectDn, "Error");
                return null;
            }
            return strValue;
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
                usrDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, ldapDomain, log);
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
        public bool CreateUserAccount(string parentOUDN, string samName, string userPassword, string firstName, string lastName, LogFile log)
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
                log.addTrn("Failure to create AD user account " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");

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
                    //MessageBox.Show("CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + " group already exists from adding");
                    return false;
                }
            }
            catch (Exception)
            {
                //MessageBox.Show(ex.Message.ToString() + "issue create user LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + "\n" + ex.StackTrace.ToString());
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
        public void NicknameUpdate(AppsService service, SqlDataReader nicknames, string usernameFieldName, string userNicknameField, LogFile log)
        {
            // updates nicknames in gmail from sql data reader containing user ids and nicknames
            // finds the user with the ID listed and returns all the nicknames for that user
            // iterates trhrough the nicknames and delets all those which do not mathch the nickname from the usernickname field
            // if the nickname from the usernickname field does not exist it will be created fro the user listed
            //
            // usernameFieldName    userNicknameField
            // ------------------------------------
            // ID               |   Nickname
            // ------------------------------------
            // SB289312         |   test.user
            //
            // 

            int i = 0;
            int nicknamecount = 0;
            bool foundnickname = false;
            try
            {
                while (nicknames.Read())
                {
                    // get all nicknames for user who has the wrong nickname
                    NicknameFeed userNicknames;
                    userNicknames = service.RetrieveNicknames(nicknames[usernameFieldName].ToString());


                    // get the count so we can iterate over them\

                    nicknamecount = userNicknames.Entries.Count;
                    // iterate and delete all nicknames that are not equal to the correct nickname
                    for (i = 0; i < nicknamecount; i++)
                    {
                        try
                        {

                            NicknameEntry nicknameEntry = userNicknames.Entries[nicknamecount] as NicknameEntry;
                            if (nicknameEntry.Nickname.Name.ToString() == nicknames[userNicknameField].ToString())
                            {
                                foundnickname = true;
                            }
                            else
                            {
                                service.DeleteNickname(nicknameEntry.Nickname.Name.ToString());
                                log.addTrn("Deleting user nickname " + nicknameEntry.Nickname.Name.ToString(), "Transaction");
                            }
                        }
                        catch
                        {
                            log.addTrn("Error deleting user nickname " + nicknames[userNicknameField].ToString(), "Error");
                        }
                    }
                    // if the nickname is not found create the new nickname
                    if (foundnickname == false)
                    {
                        try
                        {
                            service.CreateNickname(nicknames[usernameFieldName].ToString(), nicknames[userNicknameField].ToString());
                            log.addTrn("Creating user nickname " + nicknames[userNicknameField].ToString() + " for user " + nicknames[usernameFieldName].ToString(), "Transaction");
                        }
                        catch
                        {
                            log.addTrn("Error adding user nickname " + nicknames[userNicknameField].ToString() + " for user " + nicknames[usernameFieldName].ToString(), "Error");
                        }
                    }
                    // reset all variables
                    foundnickname = false;
                    i = 0;
                    nicknamecount = 0;
                }
            }
            catch (Exception ex)
            {
                log.addTrn("Issue updating nicknames datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
        }
    }


    //Logic code for execution of the synchronization
    public class ObjectADSqlsyncGroup
    {
        public void ExecuteGroupSync(GroupSynch groupsyn, ConfigSettings settingsConfig, ToolSet tools, LogFile log)
        {

            StopWatch time = new StopWatch();

            string groupapp = groupsyn.Group_Append;
            string groupOU = groupsyn.BaseGroupOU;
            string sAMAccountName = "";
            string description = "";
            int count = 0;
            string sqlgroupsTable = "#FHC_GROUPS_SQLgroupsTable";
            string adGroupsTable = "#FHC_GROUPS_ADgroupsTable";
            if (settingsConfig.TempTables == true)
            {
                sqlgroupsTable = "#FHC_GROUPS_SQLgroupsTable";
                adGroupsTable = "#FHC_GROUPS_ADgroupsTable";
            }
            else
            {
                sqlgroupsTable = "FHC_GROUPS_SQLgroupsTable" + groupsyn.Group_Append;
                adGroupsTable = "FHC_GROUPS_ADgroupsTable" + groupsyn.Group_Append;
            }
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
            SqlConnection sqlConn = new SqlConnection("Data Source=" + groupsyn.DataServer + ";Initial Catalog=" + groupsyn.DBCatalog + ";Integrated Security=SSPI;Connect Timeout=360");



            sqlConn.Open();
            // Setup the OU for the program
            log.addTrn("Setup OU for the groups", "Info");
            tools.CreateOURecursive("OU=" + groupapp + "," + groupOU, log);

            // A little house cleaning to empty out tables
            
            if (settingsConfig.TempTables == false)
            {
                log.addTrn("Clear out tables for use", "Info");
                tools.DropTable(adGroupsTable, sqlConn, log);
                tools.DropTable(sqlgroupsTable, sqlConn, log);
            }

            // grab list of groups from SQL insert into a temp table
            log.addTrn("Get groups from SQL", "Info");
            SqlCommand sqlComm = new SqlCommand();
            if (groupsyn.Group_where == "")
            {
                sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + groupsyn.Group_sAMAccount + ") AS " + groupsyn.Group_sAMAccount + ", RTRIM(" + groupsyn.Group_CN + ") + '" + groupapp + "' AS " + groupsyn.Group_CN + " INTO " + sqlgroupsTable + " FROM " + groupsyn.Group_dbTable + " WHERE " + groupsyn.Group_sAMAccount + " IS NOT NULL ORDER BY " + groupsyn.Group_CN, sqlConn);
            }
            else
            {
                sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + groupsyn.Group_sAMAccount + ") AS " + groupsyn.Group_sAMAccount + ", RTRIM(" + groupsyn.Group_CN + ") + '" + groupapp + "' AS " + groupsyn.Group_CN + " INTO " + sqlgroupsTable + " FROM " + groupsyn.Group_dbTable + " WHERE " + groupsyn.Group_sAMAccount + " IS NOT NULL AND " + groupsyn.Group_where + " ORDER BY " + groupsyn.Group_CN, sqlConn);
            }


            try
            {
                sqlComm.CommandTimeout = 360;
                sqlComm.ExecuteNonQuery();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                throw;
            }



            // generate a list of fields to ask from AD
            adUpdateKeys.Add("description");
            adUpdateKeys.Add("CN");


            // grab groups from AD
            log.addTrn("Get groups from AD", "Info");
            groupsDataTable = tools.EnumerateGroupsInOUDataTable("OU=" + groupapp + "," + groupOU, adUpdateKeys, adGroupsTable, log);

            // insert groups from AD into a temp table
            if (groupsDataTable.Rows.Count > 0)
            {

                groupsTable = tools.Create_Table(groupsDataTable, adGroupsTable, sqlConn, log);


                //Find groups that we need to create
                log.addTrn("Query to find groups that need to be created", "Info");
                add = tools.QueryNotExists(sqlgroupsTable, groupsTable, sqlConn, groupsyn.Group_CN, adUpdateKeys[1].ToString(), log);



                // Create groups
                log.addTrn("Creating groups", "Info");
                while (add.Read())
                {
                    //i++;
                    sAMAccountName = (string)add[1].ToString().Trim();
                    description = (string)add[0].ToString().Trim();
                    groupObject.Add("sAMAccountName", sAMAccountName);
                    groupObject.Add("CN", sAMAccountName);
                    groupObject.Add("description", description);
                    tools.CreateGroup("OU=" + groupapp + "," + groupOU, groupObject, log);
                    groupObject.Clear();
                }

                add.Close();


                //time.Start();
                log.addTrn("Query to find groups to delete", "Info");
                delete = tools.QueryNotExists(groupsTable, sqlgroupsTable, sqlConn, adUpdateKeys[1].ToString(), groupsyn.Group_CN, log);
                // delete groups in AD
                // i = 0;
                log.addTrn("Deleting groups", "Info");
                while (delete.Read())
                {
                    tools.DeleteGroup("OU=" + groupapp + "," + groupOU, (string)delete[adUpdateKeys[1].ToString()].ToString().Trim(), log);
                }
                delete.Close();


                // Get columns from sqlgroupsTable temp table in database get columns deprcated in favor of manual building due to cannot figure out how to get the columns of a temporary table
                // SQLupdateKeys = tools.GetColumns(groupsyn.DataServer, groupsyn.DBCatalog, sqlgroupsTable);
                // make the list of fields for the sql to check when updating note these fields must be in the same order as the AD update keys
                sqlUpdateKeys.Add(groupsyn.Group_sAMAccount);
                sqlUpdateKeys.Add(groupsyn.Group_CN);

                // update assumes the both ADupdateKeys and SQLupdateKeys have the same fields, listed in the same order check  call to EnumerateGroupsInOU if this is wrong should be sAMAccountName, CN matching the SQL order
                log.addTrn("Query to find groups which need to be updated", "Info");
                update = tools.CheckUpdate(sqlgroupsTable, groupsTable, groupsyn.Group_CN, adUpdateKeys[1].ToString(), sqlUpdateKeys, adUpdateKeys, sqlConn, log);




                // update groups in ad
                // last record which matches the primary key is the one which gets inserted into the database
                log.addTrn("Updating groups", "Info");
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
                        // log.addTrn("Group update ; " + sAMAccountName + ",OU=" + groupapp + "," + groupOU + ";" + description);
                    }
                    else
                    {
                        // find it its on the server somewhere we will log the exception
                        groupDN = tools.GetObjectDistinguishedName(objectClass.group, returnType.distinguishedName, groupObject["CN"], dc, log);
                        // what if user is disabled will user mapping handle it?
                        // groups needs to be moved and updated
                        // tools.MoveADObject(groupDN, "LDAP://OU=" + groupapp + ',' + groupOU);
                        // tools.UpdateGroup("OU=" + groupapp + "," + groupOU, groupObject);
                        log.addTrn("Group cannot be updated user probabally should be in ; " + "OU=" + groupapp + "," + groupOU + " ; but was found in ; " + groupDN, "Error");
                    }
                    groupObject.Clear();
                }
                update.Close();
            }
            // we didn't find any records in AD so there is no need for the Update or delete logic to run
            else
            {
                log.addTrn("Query to get list of groups to add", "Info");
                sqlComm = new SqlCommand("SELECT * FROM " + sqlgroupsTable, sqlConn);
                try
                {
                    sqlComm.CommandTimeout = 360;
                    add = sqlComm.ExecuteReader();
                    log.addTrn(sqlComm.CommandText.ToString(), "Query");
                    log.addTrn("Adding groups", "Info");
                    while (add.Read())
                    {
                        //i++;
                        groupObject.Add("sAMAccountName", (string)add[1]);
                        groupObject.Add("CN", (string)add[1]);
                        groupObject.Add("description", (string)add[0]);
                        tools.CreateGroup("OU=" + groupapp + "," + groupOU, groupObject, log);
                        groupObject.Clear();
                    }
                    add.Close();
                }
                catch (Exception ex)
                {
                    log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                }

           }

            // users section
            string sqlgroupMembersTable = "#FHC_GROUPS_sqlusersTable";
            string ADgroupMembersTable = "#FHC_GROUPS_ADusersTable";
            if (settingsConfig.TempTables == true)
            {
                sqlgroupMembersTable = "#FCH_GROUPS_sqlusersTable";
                ADgroupMembersTable = "#FHC_GROUPS_ADusersTable";
            }
            else
            {
                sqlgroupMembersTable = "FHC_GROUPS_sqlusersTable";
                ADgroupMembersTable = "FHC_GROUPS_ADusersTable";
            }
            SqlDataReader sqlgroups;
            DataTable ADusers = new DataTable();


            // A little house cleaning to empty out tables
            log.addTrn("Clear out tables for usage", "Info");
            if (settingsConfig.TempTables == false)
            {
                tools.DropTable(sqlgroupMembersTable, sqlConn, log);
                tools.DropTable(ADgroupMembersTable, sqlConn, log);
            }

            // grab users data from sql
            log.addTrn("Get users from SQL", "Info");
            if (groupsyn.User_where == "")
            {
                sqlComm = new SqlCommand("SELECT DISTINCT 'CN=' + RTRIM(" + groupsyn.User_sAMAccount + ") + '," + groupsyn.BaseUserOU + "' AS " + groupsyn.User_sAMAccount + ", RTRIM(" + groupsyn.User_Group_Reference + ") + '" + groupapp + "' AS " + groupsyn.User_Group_Reference + " INTO " + sqlgroupMembersTable + " FROM " + groupsyn.User_dbTable, sqlConn);
            }
            else
            {
                sqlComm = new SqlCommand("SELECT DISTINCT 'CN=' + RTRIM(" + groupsyn.User_sAMAccount + ") + '," + groupsyn.BaseUserOU + "' AS " + groupsyn.User_sAMAccount + ", RTRIM(" + groupsyn.User_Group_Reference + ") + '" + groupapp + "' AS " + groupsyn.User_Group_Reference + " INTO " + sqlgroupMembersTable + " FROM " + groupsyn.User_dbTable + " WHERE " + groupsyn.User_where, sqlConn);
            }
            try
            {
                sqlComm.CommandTimeout = 360;
                sqlComm.ExecuteNonQuery();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }



            // populate datatable with users from AD groups by looping thru the list of groups from SQL and loading the cross referenced AD group members
            log.addTrn("Get users from AD by looping through groups from SQL and appending results to a table", "Info");
            sqlComm = new SqlCommand("SELECT " + groupsyn.Group_CN + " FROM " + sqlgroupsTable, sqlConn);
            try
            {
                ArrayList sqlgroupsStr = new ArrayList();
                sqlComm.CommandTimeout = 360;
                sqlgroups = sqlComm.ExecuteReader();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");
                DataTable currentOu = new DataTable();
                while (sqlgroups.Read())
                {
                    sqlgroupsStr.Add((string)sqlgroups[0]);
                }
                sqlgroups.Close();
                count = 0;
                foreach (string key in sqlgroupsStr)
                {
                    ADusers = tools.EnumerateUsersInGroupDataTable(key, ",OU=" + groupapp + "," + groupOU, groupsyn.User_sAMAccount, groupsyn.User_Group_Reference, ADgroupMembersTable, log);
                    if (count == 0)
                    {
                        // make the temp table for ou comparisons the datatable must have somethign in it to make it
                        tools.Create_Table(ADusers, ADgroupMembersTable, sqlConn, log);
                    }
                    else
                    {
                        // table is already made now we need to only add to it
                        tools.Append_to_Table(ADusers, ADgroupMembersTable, sqlConn, log);
                    }
                    count++;
                }
                // hopefully merge acts as an append
                // ADusers.Merge(tools.EnumerateUsersInGroupDataTable((string)sqlgroups[0], ",OU=" + groupapp + "," + groupOU, groupsyn.User_sAMAccount, groupsyn.User_Group_Reference, ADgroupMembersTable, log));
                // currentOu = tools.EnumerateUsersInGroupDataTable(  , ",OU=" + groupapp + "," + groupOU, groupsyn.User_sAMAccount, groupsyn.User_Group_Reference, ADgroupMembersTable, log);
                // tools.Append_to_Table(currentOu, ADgroupMembersTable, sqlConn, log);
            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }


            // compare and add/remove
            log.addTrn("Query to find the users to add ", "Info");
            add = tools.QueryNotExists(sqlgroupMembersTable, ADgroupMembersTable, sqlConn, groupsyn.User_sAMAccount, ADusers.Columns[0].ColumnName, log);
            try
            {
                while (add.Read())
                {
                    tools.AddUserToGroup((string)add[0], "CN=" + (string)add[1] + ",OU=" + groupapp + "," + groupOU, false, dc, log);
                }
            }
            catch (Exception ex)
            {
                log.addTrn("Issue adding group datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
            }
            groupObject.Clear();
            add.Close();
            SqlCommand sqlComm2 = new SqlCommand();
            string recordCount = "";
            sqlComm2 = new SqlCommand("select count(" + groupsyn.Group_CN + ") FROM " + sqlgroupMembersTable, sqlConn);
            sqlComm2.CommandTimeout = 360;
            recordCount = sqlComm2.ExecuteScalar().ToString();
            sqlComm2.Dispose();

            if (recordCount != "0")
            {
                // setup update keys
                adUpdateKeys.Clear();
                sqlUpdateKeys.Clear();
                adUpdateKeys.Add(ADusers.Columns[1].ColumnName);
                adUpdateKeys.Add(ADusers.Columns[0].ColumnName);
                sqlUpdateKeys.Add(ADusers.Columns[1].ColumnName);
                sqlUpdateKeys.Add(ADusers.Columns[0].ColumnName);

                // get list of keys which have differed. We will delete them and then next time they will be readded as the correct key\
                // users which need to be updated just get deleted and recreadted later where they need to be
                log.addTrn("Query to see which users need to be deleted", "Info");
                delete = tools.CheckUpdate( sqlgroupMembersTable, ADgroupMembersTable, groupsyn.User_sAMAccount, ADusers.Columns[0].ColumnName, sqlUpdateKeys, adUpdateKeys, sqlConn, log);
                // delete = tools.QueryNotExists(ADgroupMembersTable, sqlgroupMembersTable, sqlConn, ADusers.Columns[1].ColumnName, groupsyn.User_Group_Reference, log);
                // delete groups in AD
                log.addTrn("Deleteing users", "Info");
                while (delete.Read())
                {
                    tools.RemoveUserFromGroup((string)delete[1], "CN=" + (string)delete[0] + ",OU=" + groupapp + "," + groupOU, log);
                }
                delete.Close();
            }
            sqlConn.Close();
        }
        public void ExecuteUserSync(UserSynch usersyn, ConfigSettings settingsConfig, ToolSet tools, LogFile log)
        {
            int i;
            ArrayList debugList = new ArrayList();
            StopWatch time = new StopWatch();


            string baseOU = usersyn.BaseUserOU;
            string DC = baseOU.Substring(baseOU.IndexOf("DC"));
            string sqlForCustomFields = "";

            // Table string place holders
            string sqlUsersTable = "#FHC_USERS_SQLusersTable";
            string adUsersTable = "#FHC_USERS_ADusersTable";


            SqlDataReader add;
            SqlDataReader delete;
            SqlDataReader update;

            SearchScope scope = SearchScope.OneLevel;
            ArrayList completeSqlKeys = new ArrayList();
            ArrayList completeADKeys = new ArrayList();
            ArrayList adUpdateKeys = new ArrayList();
            ArrayList sqlUpdateKeys = new ArrayList();
            ArrayList extraFieldsToReturn = new ArrayList();
            ArrayList fields = new ArrayList();
            Dictionary<string, string> userObject = new Dictionary<string, string>();
            SqlConnection sqlConn = new SqlConnection("Data Source=" + usersyn.DataServer + ";Initial Catalog=" + usersyn.DBCatalog + ";Integrated Security=SSPI;Connect Timeout=360");


            if (settingsConfig.TempTables == true)
            {
                sqlUsersTable = "#FHC_USERS_SQLusersTable";
                adUsersTable = "#FHC_USERS_ADusersTable";
            }
            else
            {
                sqlUsersTable = "FHC_USERS_SQLusersTable";
                adUsersTable = "FHC_USERS_ADusersTable";

            }



            //SqlDataReader sqlusers;
            SqlCommand sqlComm;
            SqlCommand sqlComm2;
            string recordCount = "";
            DataTable adUsers = new DataTable();



            sqlConn.Open();
            //housecleaning
            log.addTrn("Cleaning out tables", "Info");
            if (settingsConfig.TempTables == false)
            {
                tools.DropTable(sqlUsersTable, sqlConn, log);
                tools.DropTable(adUsersTable, sqlConn, log);
            }

            //if were only updating it doesnt matter where we want ot put new users
            if (usersyn.UpdateOnly == false)
            {
                log.addTrn("Initial setup of OUs and Groups", "Info");
                // create initial ou's; will log a warning out if they already exist
                tools.CreateOURecursive(usersyn.BaseUserOU, log);
                tools.CreateOURecursive(usersyn.UserHoldingTank, log);


                // setup extentions for the user accounts to go in to the right ou's
                userObject.Add("sAMAccountName", usersyn.UniversalGroup.Remove(0, 3).Remove(usersyn.UniversalGroup.IndexOf(",") - 3));
                userObject.Add("CN", usersyn.UniversalGroup.Remove(0, 3).Remove(usersyn.UniversalGroup.IndexOf(",") - 3));
                userObject.Add("description", "Universal Group For Users");
                // creates the group if it does not exist
                tools.CreateGroup(usersyn.UniversalGroup.Remove(0, usersyn.UniversalGroup.IndexOf(",") + 1), userObject, log);
            }


            // need to add this field first to use as a primary key when checking for existance in AD
            completeSqlKeys.Add("sAMAccountName");
            completeSqlKeys.Add("CN");
            completeSqlKeys.Add("sn");
            completeSqlKeys.Add("givenName");
            completeSqlKeys.Add("homePhone");
            completeSqlKeys.Add("st");
            completeSqlKeys.Add("streetAddress");
            completeSqlKeys.Add("l");
            completeSqlKeys.Add("postalCode");
            // ?????? MIGHT NOT BE USED


            // Lets make the SQL fields to check for update
            sqlUpdateKeys.Add("sn");
            sqlUpdateKeys.Add("givenName");
            sqlUpdateKeys.Add("homePhone");
            sqlUpdateKeys.Add("st");
            sqlUpdateKeys.Add("streetAddress");
            sqlUpdateKeys.Add("l");
            sqlUpdateKeys.Add("postalCode");



            // Lets make the Active Directory Keys as well
            completeADKeys.Add("sAMAccountName");
            completeADKeys.Add("CN");
            completeADKeys.Add("sn");
            completeADKeys.Add("givenName");
            completeADKeys.Add("homePhone");
            completeADKeys.Add("st");
            completeADKeys.Add("streetAddress");
            completeADKeys.Add("l");
            completeADKeys.Add("postalCode");
            completeADKeys.Add("distinguishedName");

            // Lets make the Active Directory fields to check for update
            adUpdateKeys.Add("sn");
            adUpdateKeys.Add("givenName");
            adUpdateKeys.Add("homePhone");
            adUpdateKeys.Add("st");
            adUpdateKeys.Add("streetAddress");
            adUpdateKeys.Add("l");
            adUpdateKeys.Add("postalCode");

            //build custom keys
            for (i = 0; i < usersyn.UserCustoms.Rows.Count; i++)
            {
                // build keys to pull back from SQL
                // as well keys to check if these fields need updating
                completeSqlKeys.Add(usersyn.UserCustoms.Rows[i][0].ToString());
                sqlUpdateKeys.Add(usersyn.UserCustoms.Rows[i][0].ToString());

                // build keys to pull back from AD
                // as well keys to check if these fields need updating
                completeADKeys.Add(usersyn.UserCustoms.Rows[i][0].ToString());
                adUpdateKeys.Add(usersyn.UserCustoms.Rows[i][0].ToString());

                // build fields to pull back from SQL
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
            log.addTrn("Get users from SQL tables", "Info");
            if (usersyn.User_where == "")
            {
                sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + usersyn.User_sAMAccount + ") AS sAMAccountName" +
                    ", RTRIM(" + usersyn.User_CN + ") AS CN" +
                    ", RTRIM(" + usersyn.User_Lname + ") AS sn" +
                    ", RTRIM(" + usersyn.User_Fname + ") AS givenName" +
                    ", RTRIM(" + usersyn.User_Mobile + ") AS homePhone" +
                    ", RTRIM(" + usersyn.User_State + ") AS st" +
                    ", RTRIM(" + usersyn.User_Address + ") AS streetAddress" +
                    //", RTRIM(" + usersyn.User_mail + ") AS mail" +
                    ", RTRIM(" + usersyn.User_city + ") AS l" +
                    ", RTRIM(" + usersyn.User_Zip + ") AS postalCode" +
                    ", RTRIM(" + usersyn.User_password + ") AS password" +
                    sqlForCustomFields +
                    " INTO " + sqlUsersTable + " FROM " + usersyn.User_dbTable, sqlConn);
            }
            else
            {
                sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + usersyn.User_sAMAccount + ") AS sAMAccountName" +
                    ", RTRIM(" + usersyn.User_CN + ") AS CN" +
                    ", RTRIM(" + usersyn.User_Lname + ") AS sn" +
                    ", RTRIM(" + usersyn.User_Fname + ") AS givenName" +
                    ", RTRIM(" + usersyn.User_Mobile + ") AS homePhone" +
                    ", RTRIM(" + usersyn.User_State + ") AS st" +
                    ", RTRIM(" + usersyn.User_Address + ") AS streetAddress" +
                    //", RTRIM(" + usersyn.User_mail + ") AS mail" +
                    ", RTRIM(" + usersyn.User_city + ") AS l" +
                    ", RTRIM(" + usersyn.User_Zip + ") AS postalCode" +
                    ", RTRIM(" + usersyn.User_password + ") AS password" +
                    sqlForCustomFields +
                    " INTO " + sqlUsersTable + " FROM " + usersyn.User_dbTable +
                    " WHERE " + usersyn.User_where, sqlConn);
            }
            try
            {
                sqlComm.CommandTimeout = 360;
                sqlComm.ExecuteNonQuery();
                log.addTrn(sqlComm.CommandText.ToString(), "Query");

            }
            catch (Exception ex)
            {
                log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                throw;
            }
            if (usersyn.SearchScope == "Subtree")
            {
                scope = SearchScope.Subtree;
            }
                    
            // go grab all the users from AD
            log.addTrn("Get users from active directory", "Info");
            adUsers = tools.EnumerateUsersInOUDataTable(usersyn.BaseUserOU, completeADKeys, adUsersTable, scope, log);
            if (adUsers.Rows.Count > 0)
            {
                // make the temp table for ou comparisons

                tools.Create_Table(adUsers, adUsersTable, sqlConn, log);

                // Quick check to stop adding if the update only box is checked
                if (usersyn.UpdateOnly == false)
                {
                    // compare query for the add/remove
                    log.addTrn("Query to find users to add", "Info");
                    add = tools.QueryNotExists(sqlUsersTable, adUsersTable, sqlConn, "sAMAccountName", adUsers.Columns[0].ColumnName, log);

                    // actual add stuff
                    log.addTrn("Adding users", "Info");
                    tools.CreateUsersAccounts(usersyn.UserHoldingTank, add, usersyn.UniversalGroup, DC, usersyn, log);
                    add.Close();

                    sqlComm2 = new SqlCommand("select count(sAMAccountName) FROM " + sqlUsersTable, sqlConn);
                    sqlComm2.CommandTimeout = 360;
                    recordCount = sqlComm2.ExecuteScalar().ToString();
                    sqlComm2.Dispose();

                    if (recordCount != "0")
                    {
                        // compare query to find records which need deletion
                        log.addTrn("Query to find users to delete", "Info");
                        delete = tools.QueryNotExists(adUsersTable, sqlUsersTable, sqlConn, usersyn.User_sAMAccount, completeADKeys[0].ToString(), log);

                        // delete users in AD
                        log.addTrn("Deleting users", "Info");
                        try
                        {
                            while (delete.Read())
                            {

                                tools.DeleteUserAccount((string)delete["distinguishedname"], log);
                                // log.addTrn("User removed ;" + (string)delete[adUpdateKeys[1].ToString()].ToString().Trim()); 
                            }
                        }
                        catch (Exception ex)
                        {
                            log.addTrn("Issue deleting AD users datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                        }
                        delete.Close();
                    }
                }

                // add the extra fields in form ".field ,"
                extraFieldsToReturn.Add(adUsersTable + ".distinguishedname ,");

                log.addTrn("Query to find users to update", "Info");
                update = tools.CheckUpdate(sqlUsersTable, adUsersTable, "sAMAccountName", "sAMAccountName", sqlUpdateKeys, adUpdateKeys, extraFieldsToReturn, 1, sqlConn, log);

                // update users in ad
                // last record which matches the primary key is the one which gets inserted into the database
                log.addTrn("Updating users", "Info");
                tools.UpdateUsers(update, DC, usersyn, log);

                update.Close();
            }
            // did not find any records in AD we are only adding users
            else
            {
                // and we are not updating users
                if (usersyn.UpdateOnly == false)
                {
                    // add the users without doing additional checks
                    tools.Create_Table(adUsers, adUsersTable, sqlConn, log);
                    log.addTrn("Query to find users to add", "Info");
                    add = tools.QueryNotExists(sqlUsersTable, adUsersTable, sqlConn, "sAMAccountName", adUsers.Columns[0].ColumnName, log);
                    log.addTrn("Add all users", "Info");
                    tools.CreateUsersAccounts(usersyn.UserHoldingTank, add, usersyn.UniversalGroup, DC, usersyn, log);
                    add.Close();
                }
            }
            sqlConn.Close();
        }
    }
    public class ObjectADGoogleSync
    {
        public void EmailUsersSync(GmailUsers gusersyn, ConfigSettings settingsConfig, ToolSet tools, LogFile log)
        {
            // MessageBox.Show("gmail " + gusersyn.Admin_password + " " + gusersyn.Admin_domain + " " + gusersyn.Admin_user + " " + gusersyn.DataServer + " " + gusersyn.DBCatalog + " " + gusersyn.User_ad_OU + " " + gusersyn.User_Datasource + " " + gusersyn.User_dbTable + " " + gusersyn.User_Fname + " " + gusersyn.User_Lname + " " + gusersyn.User_Mname + " " + gusersyn.User_password + " " + gusersyn.User_password_short_fix_checkbox.ToString() + " " + gusersyn.User_password_generate_checkbox.ToString() + " " + gusersyn.User_StuID + " " + gusersyn.User_table_view + " " + gusersyn.User_where + " " + gusersyn.Writeback_AD_checkbox.ToString() + " " + gusersyn.Writeback_ad_OU + " " + gusersyn.Writeback_DB_checkbox.ToString() + " " + gusersyn.Writeback_email_field + " " + gusersyn.Writeback_primary_key + " " + gusersyn.Writeback_secondary_email_field + " " + gusersyn.Writeback_table + " " + gusersyn.Writeback_transfer_email_checkbox.ToString() + " " + gusersyn.Writeback_where_clause);
            // Email addresses are static so only the names can be updated. passwords will be ignored
            // appservice variables will come from a config designed ot hold its data (sql and Gmail login)
            string userDN = "";
            AppsService service = new AppsService(gusersyn.Admin_domain, gusersyn.Admin_user + "@" + gusersyn.Admin_domain, gusersyn.Admin_password);
            ArrayList completeSqlKeys = new ArrayList();
            ArrayList completeGmailKeys = new ArrayList();
            ArrayList gmailUpdateKeys = new ArrayList();
            ArrayList sqlUpdateKeys = new ArrayList();
            ArrayList adUpdateKeys = new ArrayList();
            ArrayList additionalKeys = new ArrayList();

            // Table place holders
            string sqlUsersTable = "#FHC_LDAP_sqlusersTable";
            string gmailUsersTable = "#FHC_LDAP_gmailusersTable";
            string nicknamesFromGmailTable = "#FHC_LDAP_gmailNicknamesTable";
            string loginWithoutNicknamesTable = "#FHC_LDAP_loginsWONicknamesTable";
            string adNicknamesTable = "#FHC_LDAP_adNicknamesTable";
            string sqlNicknamesTable = "#FHC_LDAP_sqlNicknamesTable";
            string nicknamesToUpdateDBTable = "#FHC_LDAP_nicknamesToUpdateDB";
            string nicknamesFilteredForDuplicatesTable = "#FHC_LDAP_nicknamesFilteredDuplicates";
            string nicknamesFromGmailTable2 = "#FHC_LDAP_gmailNicknamesTable2";
            string gmailUsersTableWB = "#FHC_LDAP_gmailusersTableWB";


            SqlDataReader add;
            //SqlDataReader delete;
            SqlDataReader update;
            SqlConnection sqlConn = new SqlConnection("Data Source=" + gusersyn.DataServer + ";Initial Catalog=" + gusersyn.DBCatalog + ";Integrated Security=SSPI;Connect Timeout=360;");


            if (settingsConfig.TempTables == true)
            {
                sqlUsersTable = "#FHC_LDAP_sqlusersTable";
                gmailUsersTable = "#FHC_LDAP_gmailusersTable";
            }
            else
            {
                sqlUsersTable = "FHC_LDAP_sqlusersTable";
                gmailUsersTable = "FHC_LDAP_gmailusersTable";
            }


            //SqlDataReader sqlusers;
            SqlCommand sqlComm;
            DataTable gmailUsers = new DataTable();
            DataTable adUsers = new DataTable();
            DataTable writeback = new DataTable();


            // set up fields to pull back from SQL or AD if the flag is checked both must contain the same data
            completeSqlKeys.Add(gusersyn.User_StuID);
            completeSqlKeys.Add(gusersyn.User_Fname);
            completeSqlKeys.Add(gusersyn.User_Lname);
            completeSqlKeys.Add(gusersyn.User_Mname);
            completeSqlKeys.Add(gusersyn.User_password);

            // Lets make the SQL fields to check for update
            sqlUpdateKeys.Add(gusersyn.User_Fname);
            sqlUpdateKeys.Add(gusersyn.User_Lname);


            // Lets make the gmail Keys as well
            completeGmailKeys.Add(gusersyn.User_StuID);
            completeGmailKeys.Add(gusersyn.User_Fname);
            completeGmailKeys.Add(gusersyn.User_Mname);
            completeGmailKeys.Add(gusersyn.User_Lname);
            completeGmailKeys.Add(gusersyn.User_password);


            // Lets make the gmail fields to check for update
            gmailUpdateKeys.Add(gusersyn.User_Fname);
            gmailUpdateKeys.Add(gusersyn.User_Lname);

            // List of extra fields to pull when dealing with accounts for update since we only compare first and last name we need to repull ID and middle name
            additionalKeys.Add(sqlUsersTable + "." + gusersyn.User_StuID + ", ");
            additionalKeys.Add(sqlUsersTable + "." + gusersyn.User_Mname + ", ");

            sqlConn.Open();

            //housecleaning
            log.addTrn("Clear out tables for use", "Info");
            if (settingsConfig.TempTables == false)
            {
                tools.DropTable(sqlUsersTable, sqlConn, log);
                tools.DropTable(gmailUsersTable, sqlConn, log);
            }

            // this statement picks the datasource SQL vs AD and sets up the temp table
            
            if (gusersyn.User_Datasource == "database")
            {
                // grab users data from sql
                log.addTrn("Get users from SQL", "Info");
                if (gusersyn.User_where == "")
                {
                    sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + gusersyn.User_StuID + ")" +
                        ", RTRIM(" + gusersyn.User_Fname + ") AS " + gusersyn.User_Fname +
                        ", RTRIM(" + gusersyn.User_Lname + ") AS " + gusersyn.User_Lname +
                        ", RTRIM(" + gusersyn.User_Mname + ") AS " + gusersyn.User_Mname +
                        ", RTRIM(" + gusersyn.User_password + ") AS " + gusersyn.User_password +
                        " INTO " + sqlUsersTable + " FROM " + gusersyn.User_dbTable, sqlConn);
                }
                else
                {
                    sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + gusersyn.User_StuID +
                        ", RTRIM(" + gusersyn.User_Fname + ")" + gusersyn.User_Fname +
                        ", RTRIM(" + gusersyn.User_Lname + ")" + gusersyn.User_Lname +
                        ", RTRIM(" + gusersyn.User_Mname + ")" + gusersyn.User_Mname +
                        ", RTRIM(" + gusersyn.User_password + ")" + gusersyn.User_password +
                        " INTO " + sqlUsersTable + " FROM " + gusersyn.User_dbTable +
                        " WHERE " + gusersyn.User_where, sqlConn);
                }
                try
                {
                    sqlComm.CommandTimeout = 360;
                    sqlComm.ExecuteNonQuery();
                    log.addTrn(sqlComm.CommandText.ToString(), "Query");
                }
                catch (Exception ex)
                {
                    log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                    // If we add deletion we need to fail out if this is blank
                    //throw;
                }
            }
            else
            {
                log.addTrn("Get users from AD", "Info");
                adUsers = tools.EnumerateUsersInOUDataTable(gusersyn.User_ad_OU, completeSqlKeys, sqlUsersTable, SearchScope.OneLevel, log);
                // build the database temp table from the users retrieved into adUsers
                tools.Create_Table(adUsers, sqlUsersTable, sqlConn, log);
            }

            // go grab all the users from Gmail from the database
            log.addTrn("Get users from Gmail for comparison", "Info");
            gmailUsers = tools.Get_Gmail_Users(service, gusersyn, gmailUsersTable, log);
            // make the temp table for ou comparisons
            tools.Create_Table(gmailUsers, gmailUsersTable, sqlConn, log);



            // compare and add/remove
            log.addTrn("Query to find the users who need to be created", "Info");
            add = tools.QueryNotExists(sqlUsersTable, gmailUsersTable, sqlConn, gusersyn.User_StuID, gmailUsers.Columns[0].ColumnName, log);

            log.addTrn("Adding users", "Info");
            tools.Create_Gmail_Users(service, gusersyn, add, log);
            add.Close();

            //delete = tools.QueryNotExists(sqlUsersTable, gmailUsersTable, sqlConn, gusersyn.User_StuID, completeGmailKeys[0].ToString());
            // delete groups in AD
            //while (delete.Read())
            //{
            //tools.DeleteGmailUserAccount((string)delete[0], (string)delete[1], log);
            //
            //}
            //delete.Close();

            log.addTrn("Query to find users to update", "Info");
            update = tools.CheckUpdate(sqlUsersTable, gmailUsersTable, gusersyn.User_StuID, gmailUsers.Columns[0].ColumnName, sqlUpdateKeys, gmailUpdateKeys, additionalKeys, 0, sqlConn, log);

            log.addTrn("Updating users", "Info");
            tools.UpdateGmailUser(service, gusersyn, update, log);
            update.Close();





            // ***********************************
            // ** Start writeback features
            // ***********************************



            if (settingsConfig.TempTables == true)
            {
                nicknamesFromGmailTable = "#FHC_LDAP_gmailNicknamesTable";
                loginWithoutNicknamesTable = "#FHC_LDAP_loginsWONicknamesTable";
                adNicknamesTable = "#FHC_LDAP_adNicknamesTable";
                sqlNicknamesTable = "#FHC_LDAP_sqlNicknamesTable";
                nicknamesToUpdateDBTable = "#FHC_LDAP_nicknamesToUpdateDB";
                nicknamesFilteredForDuplicatesTable = "#FHC_LDAP_nicknamesFilteredDuplicates";
                nicknamesFromGmailTable2 = "#FHC_LDAP_gmailNicknamesTable2";
                gmailUsersTableWB = "#FHC_LDAP_gmailusersTableWB";
            }
            else
            {
                nicknamesFromGmailTable = "FHC_LDAP_gmailNicknamesTable";
                loginWithoutNicknamesTable = "FHC_LDAP_loginsWONicknamesTable";
                adNicknamesTable = "FHC_LDAP_adNicknamesTable";
                sqlNicknamesTable = "FHC_LDAP_sqlNicknamesTable";
                nicknamesToUpdateDBTable = "FHC_LDAP_nicknamesToUpdateDB";
                nicknamesFilteredForDuplicatesTable = "FHC_LDAP_nicknamesFilteredDuplicates";
                nicknamesFromGmailTable2 = "FHC_LDAP_gmailNicknamesTable2";
                gmailUsersTableWB = "FHC_LDAP_gmailusersTableWB";
            }

            string dc = gusersyn.Writeback_ad_OU.Substring(gusersyn.Writeback_ad_OU.IndexOf("DC"));
            string userNickName = "";
            SqlDataReader nicknamesToAddToAD;
            SqlDataReader lostNicknames;
            ArrayList nicknameKeys = new ArrayList();
            ArrayList sqlkeys = new ArrayList();
            ArrayList adPullKeys = new ArrayList();
            ArrayList adMailUpdateKeys = new ArrayList();
            ArrayList nicknameKeysAndTable = new ArrayList();
            ArrayList sqlkeysAndTable = new ArrayList();
            ArrayList keywordFields = new ArrayList(); //fields for checking against nickname to see how close it is to the real data
            ArrayList selectFields = new ArrayList(); //fields from nicknamesFromGmailTable to bring back
            DataTable nicknames = new DataTable();
            SqlDataReader sendAsAliases;


            // housecleaning
            log.addTrn("Clear out tables for use", "Info");
//            if (settingsConfig.TempTables == false)
//            {
                tools.DropTable(nicknamesFromGmailTable, sqlConn, log);
                tools.DropTable(loginWithoutNicknamesTable, sqlConn, log);
                tools.DropTable(adNicknamesTable, sqlConn, log);
                tools.DropTable(sqlNicknamesTable, sqlConn, log);
                tools.DropTable(nicknamesToUpdateDBTable, sqlConn, log);
                tools.DropTable(nicknamesFilteredForDuplicatesTable, sqlConn, log);
                tools.DropTable(nicknamesFromGmailTable2, sqlConn, log);
                tools.DropTable(gmailUsersTableWB, sqlConn, log);
//            }

            // install levenstein if bulid nicknames checked
            log.addTrn("Add levenshtein", "Info");
            if (gusersyn.Levenshtein == true)
            {
                string sqlcmd = "CREATE function LEVENSHTEIN( @s varchar(50), @t varchar(50) ) \n --Returns the Levenshtein Distance between strings s1 and s2. \n " +
                              "--Original developer: Michael Gilleland    http://www.merriampark.com/ld.htm \n --Translated to TSQL by Joseph Gama \n returns varchar(50) \n " +
                              "as \n BEGIN \n DECLARE @d varchar(2500), @LD int, @m int, @n int, @i int, @j int, \n @s_i char(1), @t_j char(1),@cost int \n --Step 1 \n SET @n=LEN(@s) \n" +
                              " SET @m=LEN(@t) \n SET @d=replicate(CHAR(0),2500) \n If @n = 0 \n BEGIN \n SET @LD = @m \n GOTO done \n END \n If @m = 0 \n BEGIN \n	SET @LD = @n \n" +
                              "	GOTO done \n END \n --Step 2 \n SET @i=0 \n WHILE @i<=@n \n	BEGIN \n SET @d=STUFF(@d,@i+1,1,CHAR(@i))--d(i, 0) = i \n SET @i=@i+1 \n END \n" +
                              " SET @i=0 \n WHILE @i<=@m \n BEGIN \n SET @d=STUFF(@d,@i*(@n+1)+1,1,CHAR(@i))--d(0, j) = j \n SET @i=@i+1 \n	END \n --goto done \n --Step 3 \n" +
                              " SET @i=1 \n WHILE @i<=@n \n BEGIN \n SET @s_i=(substring(@s,@i,1)) \n --Step 4 \n SET @j=1 \n	WHILE @j<=@m \n	BEGIN \n SET @t_j=(substring(@t,@j,1)) \n" +
                              " --Step 5 \n If @s_i = @t_j \n	SET @cost=0 \n ELSE \n SET @cost=1 \n --Step 6 \n SET @d=STUFF(@d,@j*(@n+1)+@i+1,1,CHAR(dbo.MIN3( \n" +
                              " ASCII(substring(@d,@j*(@n+1)+@i-1+1,1))+1, \n ASCII(substring(@d,(@j-1)*(@n+1)+@i+1,1))+1, \n ASCII(substring(@d,(@j-1)*(@n+1)+@i-1+1,1))+@cost) \n )) \n" +
                              " SET @j=@j+1 \n END \n SET @i=@i+1 \n END \n --Step 7 \n SET @LD = ASCII(substring(@d,@n*(@m+1)+@m+1,1)) \n done: \n --RETURN @LD \n" +
                              " --I kept this code that can be used to display the matrix with all calculated values \n --From Query Analyser it provides a nice way to check the algorithm in action \n" +
                              " -- \n RETURN @LD \n --declare @z varchar(8000) \n --set @z='' \n --SET @i=0 \n --WHILE @i<=@n \n --	BEGIN \n --	SET @j=0 \n --	WHILE @j<=@m \n --		BEGIN \n" +
                              " --		set @z=@z+CONVERT(char(3),ASCII(substring(@d,@i*(@m+1 )+@j+1 ,1))) \n --		SET @j=@j+1  \n --		END \n --	SET @i=@i+1 \n --	END \n --print dbo.wrap(@z,3*(@n+1)) \n END \n";
                sqlComm = new SqlCommand(sqlcmd, sqlConn);

                try
                {
                    sqlComm.CommandTimeout = 360;
                    // sqlComm.CommandType = System.Data.CommandType.StoredProcedure;
                    sqlComm.ExecuteNonQuery();
                    log.addTrn(sqlComm.CommandText.ToString(), "Query");
                }
                catch (Exception ex)
                {
                    log.addTrn("Failed to install levenstein SQL command or levenstein already installed" + sqlComm.CommandText.ToString() + " error " + ex + "\n" + ex.StackTrace.ToString(), "Error");
                }
                // install min3
                sqlcmd = "CREATE function MIN3(@a int,@b int,@c int ) \n --Returns the smallest of 3 numbers. \n" +
                    "returns int \n as \n BEGIN \n declare @temp int \n if (@a < @b)  AND (@a < @c) \n select @temp=@a \n else \n if (@b < @a)  AND (@b < @c) \n select @temp=@b \n else \n" +
                    "select @temp=@c \n return @temp \n END";
                sqlComm = new SqlCommand(sqlcmd, sqlConn);
                try
                {
                    //sqlComm.CommandType = System.Data.CommandType.StoredProcedure;
                    sqlComm.CommandTimeout = 360;
                    sqlComm.ExecuteNonQuery();
                    log.addTrn(sqlComm.CommandText.ToString(), "Query");
                }
                catch (Exception ex)
                {
                    log.addTrn("Failed to install min3 SQL command or min3 already installed" + sqlComm.CommandText.ToString() + " error " + ex + "\n" + ex.StackTrace.ToString(), "Error");
                }
            }

            if (gusersyn.Writeback_AD_checkbox == true || gusersyn.Writeback_DB_checkbox == true)
            {
                // DATABASE writeback

                // Make preperations to pull all data into seperate tables
                // build sql to run to get gmail user nicknames
                // execute data pull for SQL nicknames used in writeback to SQL database
                log.addTrn("Pull SQL data fro writeback of emails", "Info");
                if (gusersyn.Writeback_DB_checkbox == true)
                {
                    if (gusersyn.Writeback_where_clause == "")
                    {
                        sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + gusersyn.Writeback_primary_key + ") AS " + gusersyn.Writeback_primary_key +
                            ", RTRIM(" + gusersyn.Writeback_email_field + ") AS " + gusersyn.Writeback_email_field +
                            " INTO " + sqlNicknamesTable + " FROM " + gusersyn.Writeback_table, sqlConn);
                    }
                    else
                    {
                        sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + gusersyn.Writeback_primary_key + ") AS " + gusersyn.Writeback_primary_key +
                            ", RTRIM(" + gusersyn.Writeback_email_field + ") AS " + gusersyn.Writeback_email_field +
                            " INTO " + sqlNicknamesTable + " FROM " + gusersyn.Writeback_table +
                            " WHERE " + gusersyn.Writeback_where_clause, sqlConn);
                    }
                    try
                    {
                        sqlComm.CommandTimeout = 360;
                        sqlComm.ExecuteNonQuery();
                        log.addTrn(sqlComm.CommandText.ToString(), "Query");
                    }
                    catch (Exception ex)
                    {
                        log.addTrn("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex + "\n" + ex.StackTrace.ToString(), "Error");
                    }
                }


                // TABLE pulled by command from gmail
                // sqlNicknamesTable::  #sqlNicknamesTable
                //---------------------------------------------
                // gusersyn.writebackPrimaryKey::  soc_sec           xx111111
                // gusersyn.writebackEmailField::  gmail             first.last@domain

                // TABLE pulled by command from AD
                // adNicknamesTable::  #adNicknamesTable
                //---------------------------------------------
                // sAMAccountName                                    xx111111
                // givenName                                         first
                // middleName                                        middle
                // sn                                                last
                // mail                                              first.last@domain

                // TABLE pulled by command from gamil
                // nicknamesFromGmailTable:: #nicknamesFromGmailTable
                //---------------------------------------------
                // gusersyn.Writeback_primary_key:: soc_sec          xx111111
                // nicknames.column[1]:: nickname                    first.last
                // nicknames.column[2]:: email                       first.last@domain


                // TABLE pulled by command from gmail
                // gmailUsersTable:: #gmailUsersTable
                //---------------------------------------------
                // gusersyn.User_StuID:: sAMAccountName              xx111111
                // gusersyn.Fname:: givenname                        first
                // gusersyn.Lname:: sn                               last


                // TABLE generated by sql query from #nicknamesFromGmailTable and #gmailUsersTable
                // loginWithoutNicknamesTable:: #loginWithoutNicknamesTable
                //---------------------------------------------
                // gusersyn.User_StuID:: sAMAccountName              xx111111
                // gusersyn.Fname:: givenname                        first
                // gusersyn.Lname:: sn                               last


                // TABLE created by query to database
                // sqlUsersTable:: #sqlUsersTable 
                //---------------------------------------------
                // gusersyn.User_StuID:: SQL or AD                   xx111111
                // gusersyn.User_Fname:: SQL or AD                   first
                // gusersyn.User_Lname:: SQL or AD                   middle
                // gusersyn.User_Mname:: SQL or AD                   last
                // gusersyn.User_password:: SQL or AD                *******



                // TABLE created by query to database
                // nicknamesToUpdateDBTable:: #nicknamesToUpdateDBTable 
                //---------------------------------------------
                //nicknames.Columns[0].ColumnName:: gusersyn.Writeback_primary_key :: soc_sec           xx111111
                //nicknames.Columns[2].ColumnName:: email                                               first.last@domain



                // were adding the mail field to the pull fields not so we need to repull
                // keys to pull from database pull are most likely wrong need to hard code appropriate keys
                adPullKeys.Add("sAMAccountName");
                adPullKeys.Add("givenName");
                adPullKeys.Add("middleName");
                adPullKeys.Add("sn");
                adPullKeys.Add("mail");



                // clear our previous data
                log.addTrn("Get AD user data for nickname comparison", "Info");
                adUsers.Clear();
                adUsers = tools.EnumerateUsersInOUDataTable(gusersyn.User_ad_OU, adPullKeys, adNicknamesTable, SearchScope.OneLevel, log);
                tools.Create_Table(adUsers, adNicknamesTable, sqlConn, log);




                // get list of users from gmail this may have changed when we ran the update to test to see if anyone is missing nicknames
                log.addTrn("Get Gmail user data for user comparison", "Info");
                gmailUsers.Clear();
                gmailUsers = tools.Get_Gmail_Users(service, gusersyn, gmailUsersTableWB, log);
                tools.Create_Table(gmailUsers, gmailUsersTableWB, sqlConn, log);
                if (gmailUsers.Rows.Count > 0)
                {
                    // get list of nicknames from gmail
                    log.addTrn("Get Gmail nickname data for nickname comparison", "Info");
                    nicknames.Clear();
                    nicknames = tools.Get_Gmail_Nicknames(service, gusersyn, nicknamesFromGmailTable, log);
                    tools.Create_Table(nicknames, nicknamesFromGmailTable, sqlConn, log);

                    // if we did not get any nicknames there will be problems
                    if (nicknames.Rows.Count > 0)
                    {
                        // check which user do not have a an associated nickname with them 
                        // cross reference for null userID's in nickname service.RetrieveAllNicknames table with list of all userlogin userID's from gmail service.RetrieveAllUsers
                        log.addTrn("Query to find users which do not have a nickname", "Info");
                        tools.QueryNotExistsIntoNewTable(gmailUsersTableWB, nicknamesFromGmailTable, loginWithoutNicknamesTable, sqlConn, gusersyn.User_StuID, nicknames.Columns[0].ColumnName, log);


                        // use retrieved list of users without nicknames and check for updates against list of users in main datasource
                        // we dont want ot add nicknames to people wo aren't in the primary datasource
                        // use the datatable from the view/table as the primary data source
                        // this table is generated above during the user account addition and update section
                        // sqlUsersTable will have AD or SQL no matter which we checked, there was only on variable used
                        // Pulls back the account infromation from the add/update section for users without a nickname
                        // loginWithoutNicknamesTable inherits key from from #gmailuserstable during QueryNotExistsIntoNewTable due to only pulling data from the gmailuserstable
                        log.addTrn("Query to cross reference the primary datasource to find the users who need a nickname", "Info");
                        lostNicknames = tools.QueryInnerJoin(sqlUsersTable, loginWithoutNicknamesTable, gusersyn.User_StuID, gusersyn.User_StuID, sqlConn, log);
                        // iterate lostnicknames and create nicknames
                        try
                        {
                            log.addTrn("Creating new nicknames", "Info");
                            while (lostNicknames.Read())
                            {
                                userNickName = tools.GetNewUserNickname(service, lostNicknames[gusersyn.User_StuID.ToString()].ToString(), lostNicknames[gusersyn.User_Fname.ToString()].ToString(), lostNicknames[gusersyn.User_Mname.ToString()].ToString(), lostNicknames[gusersyn.User_Lname.ToString()].ToString(), 0, false);
                                log.addTrn("Added Gmail user " + lostNicknames[gusersyn.User_StuID.ToString()].ToString() + "@" + gusersyn.Admin_domain + " Aliased as " + userNickName + "@" + gusersyn.Admin_domain, "Transaction");
                            }
                        }
                        catch (Exception ex)
                        {
                            log.addTrn("Issue creating new nicknames datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                        }

                        lostNicknames.Close();

                    }
                    // get list of nicknames from gmail it may have changed in the last update.
                    log.addTrn("Get nicknames to compare for who need new nicknames", "Info");
                    nicknames = tools.Get_Gmail_Nicknames(service, gusersyn, nicknamesFromGmailTable2, log);
                    tools.Create_Table(nicknames, nicknamesFromGmailTable2, sqlConn, log);
                    // check if we got nicknames or else it will be a problem
                    if (nicknames.Rows.Count > 0)
                    {

                        // create array lists of fields which match for updating
                        // ID field
                        nicknameKeys.Add(nicknames.Columns[0].ColumnName);
                        nicknameKeys.Add(nicknames.Columns[2].ColumnName);
                        sqlkeys.Add(gusersyn.Writeback_primary_key);
                        sqlkeys.Add(gusersyn.Writeback_email_field);

                        // check against list of nicknames in database
                        if (gusersyn.Writeback_DB_checkbox == true)
                        {

                            // check the table to see if the nicknames match the ones we have there can be more than one nickname per user  

                            //select FHC_test2_gmailNicknamesTable.soc_sec, FHC_test2_gmailNicknamesTable.Email
                            //FROM FHC_test2_gmailNicknamesTable INNER JOIN FHC_test2_sqlNicknamesTable 
                            //ON FHC_test2_gmailNicknamesTable.soc_sec = FHC_test2_sqlNicknamesTable.soc_sec 
                            //where email not in (select email from FHC_test2_gmailnicknamestable)

                            //Filter out the good nicknames


                            // gusersyn.User_Fname:: SQL or AD                   first
                            // gusersyn.User_Lname:: SQL or AD                   middle
                            // gusersyn.User_Mname:: SQL or AD                   last
                            keywordFields.Add(gusersyn.User_Lname);
                            keywordFields.Add(gusersyn.User_Fname);
                            keywordFields.Add(gusersyn.User_Mname);

                            selectFields.Add(nicknames.Columns[2].ColumnName);



                          /*  if (gusersyn.Levenshtein == true)
                            {
                                tools.SelectNicknamesClosestToActualNameIntoNewTable(nicknamesFromGmailTable, sqlUsersTable, nicknames.Columns[0].ColumnName, gusersyn.User_StuID, nicknamesFilteredForDuplicatesTable, selectFields, nicknames.Columns[1].ColumnName, keywordFields, sqlConn, log);
                                if (gusersyn.Levenshtein == true)
                                {
                                    //correct nicknames in Gmail
                                   // tools.setnames(nicknamesFromGmailTable, sqlConn, log);

                                }
                            }
                            else
                            {
                                // Filter nicknames for duplicates into new table */
                                log.addTrn("Query to find most likely nickname", "Info");
                                tools.SelectNicknamesClosestToActualNameIntoNewTable(nicknamesFromGmailTable2, sqlUsersTable, nicknames.Columns[0].ColumnName, gusersyn.User_StuID, nicknamesFilteredForDuplicatesTable, selectFields, nicknames.Columns[1].ColumnName, keywordFields, sqlConn, log);
                            // }
                            // Check filtered nicknames against the sql data to see which emails need updating and put into a table for the next operation
                            log.addTrn("Query to see if most likely nickname matches the primary datasource", "Info");
                            tools.CheckEmailUpdateIntoNewTable(nicknames.Columns[2].ColumnName, gusersyn.Writeback_email_field, nicknamesFilteredForDuplicatesTable, sqlNicknamesTable, nicknames.Columns[0].ColumnName, gusersyn.Writeback_primary_key, nicknamesToUpdateDBTable, nicknameKeys, sqlkeys, sqlConn, log);



                            // reset arraylists with just the data fields we don't to be updating the  primary key
                            nicknameKeys.Clear();
                            sqlkeys.Clear();
                            nicknameKeys.Add(nicknames.Columns[2].ColumnName);
                            sqlkeys.Add(gusersyn.Writeback_email_field);

                            adMailUpdateKeys.Add("mail");
                            if (gusersyn.Writeback_transfer_email_checkbox == true)
                            {
                                //UPDATE a1 SET a1.e_mail = a2.gmail FROM address as a1 INNER JOIN address as a2 ON a1.soc_sec = a2.soc_sec where a2.gmail <> ' '

                                nicknameKeysAndTable.Add(gusersyn.Writeback_table + "." + gusersyn.Writeback_email_field); // source table and column
                                sqlkeysAndTable.Add(gusersyn.Writeback_table + "." + gusersyn.Writeback_secondary_email_field); // target table and column

                                //UNTESTED WILL MOST LIKELY FAIL
                                log.addTrn("Query to mass move the primary email field to secondary email field in primary datasource", "Info");
                                tools.Mass_Email_Shift(gusersyn, nicknamesToUpdateDBTable, gusersyn.Writeback_table, nicknames.Columns[0].ColumnName, gusersyn.Writeback_primary_key, nicknameKeysAndTable, sqlkeysAndTable, gusersyn.Writeback_where_clause, sqlConn, log);
                            }
                            log.addTrn("Query to mass populate the email field with new aliases", "Info");
                            tools.Mass_Table_update(nicknamesToUpdateDBTable, gusersyn.Writeback_table, nicknames.Columns[0].ColumnName, gusersyn.Writeback_primary_key, nicknameKeys, sqlkeys, gusersyn.Writeback_where_clause, sqlConn, log);
                        }

                        // reset arraylists  for new update check
                        nicknameKeys.Clear();
                        nicknameKeys.Add(nicknames.Columns[0].ColumnName);
                        nicknameKeys.Add(nicknames.Columns[2].ColumnName);
                        // AD fields hard code due to not giving options in interface
                        adMailUpdateKeys.Clear();
                        adMailUpdateKeys.Add("sAMAccountName");
                        adMailUpdateKeys.Add("mail");


                        if (gusersyn.Writeback_AD_checkbox == true)
                        {
                            // find nicknames missing in AD
                            log.addTrn("Query to find AD users that need their nickname updated", "Info");
                            nicknamesToAddToAD = tools.CheckUpdate(nicknamesFilteredForDuplicatesTable, adNicknamesTable, nicknames.Columns[0].ColumnName, adPullKeys[0].ToString(), nicknameKeys, adMailUpdateKeys, sqlConn, log);
                            //nicknamesToAddToAD = tools.QueryNotExists(nicknamesFromGmailTable, adNicknamesTable, sqlConn, nicknames.Columns[0].ColumnName, adPullKeys[0].ToString(), log);
                            // interate and update mail fields
                            log.addTrn("Updating nicknames in AD", "Info");
                            try
                            {
                                while (nicknamesToAddToAD.Read())
                                {
                                    userDN = tools.GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, nicknamesToAddToAD[nicknames.Columns[0].ColumnName].ToString(), dc, log);
                                    tools.SetAttributeValuesSingleString("mail", nicknamesToAddToAD[nicknames.Columns[2].ColumnName.ToString()].ToString(), userDN, log);
                                }
                            }
                            catch (Exception ex)
                            {
                                log.addTrn("Issue writing nickname to AD datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString(), "Error");
                            }
                            nicknamesToAddToAD.Close();
                        }


                        // create send as aliases need to get data for users 

                        additionalKeys.Clear();
                        additionalKeys.Add(nicknamesFilteredForDuplicatesTable + "." + nicknames.Columns[2].ColumnName + ", ");
                        log.addTrn("Query to find new send as aliases", "Info");
                        sendAsAliases = tools.QueryInnerJoin(sqlUsersTable, nicknamesFilteredForDuplicatesTable, gusersyn.User_StuID, nicknames.Columns[0].ColumnName, additionalKeys, sqlConn, log);
                        // tools.CreateSendAs(gusersyn, sendAsAliases, nicknames.Columns[2].ColumnName, nicknames.Columns[2].ColumnName, log);

                    }
                }
            }


            sqlConn.Close();
        }
    }
}

