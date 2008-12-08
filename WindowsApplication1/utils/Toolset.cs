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
using System.Text.RegularExpressions;
using System.Web;
using System.Windows.Forms;
using Google.GData.Apps;
using Google.GData.Apps.GoogleMailSettings;
using Google.GData.Client;
using WindowsApplication1;
using WindowsApplication1.utils;


// outstanding issues
// send as use fix from google forums\
// update gmail failing to use middle name properly
// ensure nicknames are genereated properly
// allow for nulls in blank fields to be matching
// send as aliasing failing due to blank sqldata reader




namespace WindowsApplication1.utils
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
        //private String configUser_mail;
        private String configUser_table_view;
        private String configUser_dbTable;
        private String configUser_where;
        private String configDataServer;
        private String configDBCatalog;
        private String configUserHoldingTank;
        private String configUser_password;
        //private String configUserEmailDomain;
        private DataTable configCustoms = new DataTable();
        private string custom = "";
        int i = 0;
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
            //configUser_mail = "";
            configUser_where = "";
            configDataServer = "";
            configDBCatalog = "";
            configUserHoldingTank = "";
            configUser_password = "";
            //configUserEmailDomain = "";
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
            //dictionary.TryGetValue("configUser_mail", out configUser_mail);
            dictionary.TryGetValue("configUser_dbTable", out configUser_dbTable);
            dictionary.TryGetValue("configUser_where", out configUser_where);
            dictionary.TryGetValue("configDataServer", out configDataServer);
            dictionary.TryGetValue("configDBCatalog", out configDBCatalog);
            dictionary.TryGetValue("configUserHoldingTank", out configUserHoldingTank);
            dictionary.TryGetValue("configUser_password", out configUser_password);
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
            //returnvalue.Add("configUser_mail", configUser_mail);
            returnvalue.Add("configUser_dbTable", configUser_dbTable);
            returnvalue.Add("configUser_where", configUser_where);
            returnvalue.Add("configDataServer", configDataServer);
            returnvalue.Add("configDBCatalog", configDBCatalog);
            returnvalue.Add("configUniversalGroup", configUniversalGroup);
            returnvalue.Add("configUserHoldingTank", configUserHoldingTank);
            returnvalue.Add("configUser_password", configUser_password);
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



        public void load(Dictionary<string, string> dictionary)
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
            }

            catch (Exception ex)
            {
                log.errors.Add("Failure getting AD users exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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

            //string bladh = "LDAP://" + "CN=" + System.Web.HttpUtility.UrlEncode(groupDN).Replace("+", " ").Replace("*", "%2A") + groupou;
            try
            {
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
            }
            catch (Exception ex)
            {
                log.errors.Add("Failure getting AD users in a groups exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
            }
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
            }
            catch (Exception ex)
            {
                log.errors.Add("Failure getting AD groups exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                log.transactions.Add("AD set value for field " + attributeName + " for user " + objectDn + " value " + newValue);
                ent.Close();
                ent.Dispose();
            }
            catch (Exception ex)
            {
                log.errors.Add("failed AD set value for field " + attributeName + " for user " + objectDn + " value " + newValue + " exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
            }


        }
        public string GetObjectDistinguishedName(objectClass objectCls, returnType returnValue, string objectName, string LdapDomain, LogFile log)
        {
            // LdapDomain = "DC=Fabrikam,DC=COM" 

            string distinguishedName = string.Empty;
            string connectionPrefix = "LDAP://" + LdapDomain;
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
                log.errors.Add("searcher failed " + LdapDomain + " " + objectName + " Exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
            catch (Exception ex)
            {
                log.errors.Add(ex.Message.ToString() + "issue create group LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + "\n" + ex.StackTrace.ToString());
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
                    log.transactions.Add("updated group | LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath);

                }
                else
                {
                    log.warnings.Add(ouPath + " group does not exist");
                }
            }
            catch (Exception ex)
            {
                log.errors.Add(ex.Message.ToString() + "issue updating group LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + "\n" + ex.StackTrace.ToString());
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
                catch (Exception ex)
                {
                    log.errors.Add(ex.Message.ToString() + " error deleting LDAP://CN=" + name + "," + ouPath + "\n" + ex.StackTrace.ToString());
                }
            }
            else
            {
                log.warnings.Add("group LDAP://CN=" + name + "," + ouPath + " does not exists cannot delete");
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
                    log.transactions.Add("created ou | LDAP://OU=" + name + "," + ouPath);

                }
                else
                {
                    log.warnings.Add("creating ou LDAP://OU=" + name + "," + ouPath + " already exists");
                }
            }
            catch (Exception ex)
            {
                log.errors.Add(ex.Message.ToString() + "error creating ou LDAP://OU=" + name + "," + ouPath + "\n" + ex.StackTrace.ToString());
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
                    log.transactions.Add("deleting ou | LDAP://OU=" + name + "," + ouPath + " does not exists");

                }
                else
                {
                    log.warnings.Add("error deleting ou LDAP://OU=" + name + "," + ouPath + " does not exists");
                }
            }
            catch (Exception ex)
            {
                log.errors.Add(ex.Message.ToString() + " error deleting ou LDAP://OU=" + name + "," + ouPath + "\n" + ex.StackTrace.ToString());
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
            // oupath holds the path for the AD OU to hold the Users 
            // users is a sqldatareader witht the required fields in it ("CN") other Datastructures would be easy to substitute 
            // groupDN is a base group which all new users get automatically inserted into

            int i;
            int fieldcount;
            int val;
            string name = "";
            fieldcount = users.FieldCount;
            try
            {
                while (users.Read())
                {

                    if (users[usersyn.User_password].ToString() != "")
                    {
                        if (!DirectoryEntry.Exists("LDAP://CN=" + System.Web.HttpUtility.UrlEncode(users[usersyn.User_sAMAccount].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath))
                        {

                            DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);
                            DirectoryEntry newUser = entry.Children.Add("CN=" + System.Web.HttpUtility.UrlEncode(users[usersyn.User_sAMAccount].ToString()).Replace("+", " ").Replace("*", "%2A"), "user");
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
                                if (name != "password" && name != "CN")
                                {
                                    if ((string)users[i] != "")
                                    {
                                        newUser.Properties[users.GetName(i)].Value = System.Web.HttpUtility.UrlEncode((string)users[i]).Replace("+", " ").Replace("*", "%2A");
                                    }
                                }
                            }


                            AddUserToGroup("CN=" + System.Web.HttpUtility.UrlEncode(users[usersyn.User_sAMAccount].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + usersyn.UserHoldingTank, groupDn, log);
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
                    log.errors.Add("issue create user LDAP://CN=" + System.Web.HttpUtility.UrlEncode((string)users["CN"]).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + "\n" + debugdata + " failed field maybe " + name + " | " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
                    // MessageBox.Show(e.Message.ToString() + "issue create user LDAP://CN=" + System.Web.HttpUtility.UrlEncode((string)users["CN"]).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + "\n" + debugdata);
                }
                else
                {
                    log.errors.Add("issue creating users datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
        public void UpdateUsers(SqlDataReader users, LogFile log)
        {
            // requires distinguished name to be a field
            // all field names must be valid AD field names
            int fieldcount = 0;
            int i = 0;
            string name = "";
            fieldcount = users.FieldCount;
            try
            {
                while (users.Read())
                {

                    DirectoryEntry user = new DirectoryEntry("LDAP://" + (string)users["distinguishedname"]);
                    for (i = 0; i < fieldcount; i++)
                    {
                        name = users.GetName(i);
                        if (name != "password" && name != "CN" && name != "sAMAccountName" && name != "distinguishedname")
                        {
                            if ((string)users[name] != "")
                            {
                                user.Properties[name].Value = System.Web.HttpUtility.UrlEncode((string)users[name]).Replace("+", " ").Replace("*", "%2A");
                            }
                        }
                    }
                    user.CommitChanges();
                    log.transactions.Add("User updated |" + (string)users["distinguishedname"] + " ");
                }
            }
            catch (Exception ex)
            {
                if (users != null)
                {
                    log.errors.Add("issue updating user " + name + " " + System.Web.HttpUtility.UrlEncode((string)users["distinguishedname"]).Replace("+", " ").Replace("*", "%2A") + "\n" + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
                }
                else
                {
                    log.errors.Add("issue updating users data reader is null " + "\n" + ex.Message.ToString());
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
            userDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, ldapDomain, log);
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
            userDN = GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, sAMAccountName, ldapDomain, log);
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
        public bool DeleteUserAccount(string FQDN, LogFile log)
        {
            try
            {
                FQDN = "LDAP://" + FQDN;
                DirectoryEntry ent = new DirectoryEntry(FQDN);
                ent.DeleteTree();
                ent.Close();
                ent.Dispose();
                log.transactions.Add("deleted user account |" + FQDN);
                return true;
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                log.errors.Add(E.Message.ToString() + " error deleting user " + FQDN);
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
                sqlComm.ExecuteNonQuery();
                log.transactions.Add("table created " + table);
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                log.errors.Add("Failed SQL bulk copy " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                log.errors.Add("failed sql table append " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                sqlComm.ExecuteNonQuery();
                log.transactions.Add("table dropped " + table);
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
            }
        }
        // SQL query tools
        public SqlDataReader QueryNotExists(string table1, string table2, SqlConnection sqlConn, string pkey1, string pkey2, LogFile log)
        {
            // finds items in table1 who do not exist in table2 and returns the data fields table 1 for these rows
            //*************************************************************************************************
            //| Table1                      | Table2                    | Returned result
            //*************************************************************************************************
            //| ID            Data          | ID             Data       |                   | Table1.ID     Table1.DATA
            //| 1             a             | 1              a          | NOT RETURNED      |
            //| 2             b             | null           null       | RETURNED          | 3             b
            //| 3             c             | 2              null       | NOT RETURNED      |
            //| 4             d             | 4              e          | NOT RETURNED      |
            //
            // SqlCommand sqlComm = new SqlCommand("Select Table1.* Into #Table3ADTransfer From " + Table1 + " AS Table1, " + Table2 + " AS Table2 Where Table1." + pkey1 + " = Table2." + pkey2 + " And Table2." + pkey2 + " is null", sqlConn);
            SqlCommand sqlComm = new SqlCommand("SELECT DISTINCT uptoDate.* FROM " + table1 + " uptoDate LEFT OUTER JOIN " + table2 + " outofDate ON outofDate." + pkey2 + " = uptoDate." + pkey1 + " WHERE outofDate." + pkey2 + " IS NULL;", sqlConn);
            // create the command object
            SqlDataReader r;
            try
            {
                r = sqlComm.ExecuteReader();
                return r;
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
            //| 2             b             | null           null       | RETURNED          | 3             b
            //| 3             c             | 2              null       | NOT RETURNED      |
            //| 4             d             | 4              e          | NOT RETURNED      |
            // SqlCommand sqlComm = new SqlCommand("Select Table1.* Into #Table3ADTransfer From " + Table1 + " AS Table1, " + Table2 + " AS Table2 Where Table1." + pkey1 + " = Table2." + pkey2 + " And Table2." + pkey2 + " is null", sqlConn);
            SqlCommand sqlComm = new SqlCommand("SELECT DISTINCT uptoDate.* INTO " + newTable + " FROM " + table1 + " uptoDate LEFT OUTER JOIN " + table2 + " outofDate ON outofDate." + pkey2 + " = uptoDate." + pkey1 + " WHERE outofDate." + pkey2 + " IS NULL", sqlConn);
            // create the command object
            try
            {
                sqlComm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
            }


        }
        public SqlDataReader QueryInnerJoin(string table1, string table2, string pkey1, string pkey2, SqlConnection sqlConn, LogFile log)
        {
            // Returns data from table1 where the row is in both table 1 and table2
            //*************************************************************************************************
            //| Table1                      | Table2                    | Returned result
            //*************************************************************************************************
            //| ID            Data          | ID             Data       |                   | Table1.ID     Table1.DATA
            //| 1             a             | 1              a          | NOT RETURNED      |
            //| 2             b             | null           null       | NOT RETURNED      |              
            //| 3             c             | 3              null       | RETURNED          | 3             c
            //| 4             d             | 4              e          | NOT RETURNED      | 

            SqlCommand sqlComm = new SqlCommand("SELECT DISTINCT " + table1 + ".* FROM " + table1 + " INNER JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2, sqlConn);
            SqlDataReader r;
            try
            {
                r = sqlComm.ExecuteReader();
                return r;
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
            //| 1             a             | 1              a            e             | NOT RETURNED      |
            //| 2             b             | null           null         f             | NOT RETURNED      |              
            //| 3             c             | 3              null         g             | RETURNED          | 3             c               g
            //| 4             d             | 4              e            h             | NOT RETURNED      | 
            
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
                r = sqlComm.ExecuteReader();
                return r;
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
            //| 1             a             | 1              a          | NOT RETURNED      |
            //| 2             b             | null           null       | NOT RETURNED      |              
            //| 3             c             | 3              null       | RETURNED          | 3             c
            //| 4             d             | 4              e          | NOT RETURNED      | 

            SqlCommand sqlComm = new SqlCommand("SELECT DISTINCT " + table1 + ".* INTO " + newTable + " FROM " + table1 + " INNER JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2, sqlConn);
            try
            {
                sqlComm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
            }
        }
        public SqlDataReader CheckUpdate(string table1, string table2, string pkey1, string pkey2, ArrayList compareFields1, ArrayList compareFields2, SqlConnection sqlConn, LogFile log)
        {
            // NULL not handeled as blanks
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
            SqlCommand sqlComm = new SqlCommand("SELECT DISTINCT " + fields + " FROM " + table1 + " INNER JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2 + " WHERE (" + compare2 + ") <> (" + compare1 + ")", sqlConn);
            //AND " + table2 + "." + pkey2 + " != NULL
            try
            {
                SqlDataReader r = sqlComm.ExecuteReader();
                return r;
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
            }
            return null;
        }
        public SqlDataReader CheckUpdate(string table1, string table2, string pkey1, string pkey2, ArrayList compareFields1, ArrayList compareFields2, ArrayList additionalFields, SqlConnection sqlConn, LogFile log)
        {
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


            string compare1 = "";
            string compare2 = "";
            string fields = "";
            string notnull = "";
            string additionalfields = "";

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



            // need a comand builder and research on the best way to compare all fields in a row
            // this basically will just issue a concatenation sql query to the DB for each field to compare
            foreach (string key in compareFields1)
            {
                compare1 = compare1 + table1 + "." + key + " + ";
                fields += table1 + "." + key + ", ";
            }
            compare1 = compare1 + table1 + "." + pkey1 + " + ";
            foreach (string key in compareFields2)
            {
                compare2 = compare2 + table2 + "." + key + " + ";
                //fields += table2 + "." + key + ", ";
                notnull += table2 + "." + key + " <> '' OR ";
            }
            compare2 = compare2 + table2 + "." + pkey2 + " + ";
            foreach (string key in additionalFields)
            {
                additionalfields += key;
            }
            // remove trailing stuff commas +'s etc
            compare2 = compare2.Remove(compare2.Length - 2);
            compare1 = compare1.Remove(compare1.Length - 2);
            fields = fields.Remove(fields.Length - 2);
            notnull = notnull.Remove(notnull.Length - 3);
            additionalfields = additionalfields.Remove(additionalfields.Length - 2);
            SqlCommand sqlComm;
            if (additionalFields.Count > 0)
            {
                sqlComm = new SqlCommand("SELECT DISTINCT " /*+ compare2 + "," + compare1 + "," + table1 + "." + pkey1 + "," + table2 + "." + pkey2 + ","*/ + fields + ", " + additionalfields + " FROM " + table1 + " INNER JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2 + " AND (" + compare2 + ") <> (" + compare1 + ") WHERE " + notnull, sqlConn);
            }
            else
            {
                sqlComm = new SqlCommand("SELECT DISTINCT " + fields + " FROM " + table1 + " INNER JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2 + " AND (" + compare2 + ") <> (" + compare1 + ") WHERE " + notnull, sqlConn);
            }
            //AND " + table2 + "." + pkey2 + " != NULL
            try
            {
                SqlDataReader r = sqlComm.ExecuteReader();
                return r;
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
            SqlCommand sqlComm = new SqlCommand("SELECT DISTINCT " + fields + " INTO " + newTable + " FROM " + table1 + " INNER JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2 + " WHERE (" + compare2 + ") <> (" + compare1 + ")", sqlConn);
            //AND " + table2 + "." + pkey2 + " != NULL
            try
            {
                sqlComm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                sqlComm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                sqlComm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                    sqlComm.ExecuteNonQuery();
                    log.transactions.Add("DB email writeback, user " + users.Rows[i][gusersyn.Writeback_primary_key].ToString().Replace("'", "''") + ", email " + users.Rows[i][gusersyn.Writeback_email_field].ToString().Replace("'", "''"));
                }
                catch (Exception ex)
                {
                    log.errors.Add("DB email writeback failure, user " + users.Rows[i][gusersyn.Writeback_primary_key].ToString().Replace("'", "''") + ", email " + users.Rows[i][gusersyn.Writeback_email_field].ToString().Replace("'", "''") + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                        sqlComm.ExecuteNonQuery();
                        log.transactions.Add("DB email writeback, user " + users[gusersyn.Writeback_primary_key].ToString().Replace("'", "''") + ", email " + users[gusersyn.Writeback_email_field].ToString().Replace("'", "''"));
                    }
                    catch (Exception ex)
                    {
                        log.errors.Add("DB email writeback failure, user " + users[gusersyn.Writeback_primary_key].ToString().Replace("'", "''") + ", email " + users[gusersyn.Writeback_email_field].ToString().Replace("'", "''") + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
                    }
                    sqlComm.Dispose();
                }
            }
            catch (Exception ex)
            {
                log.errors.Add("Issue in DB writeback datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
            string query = "";
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
                sqlComm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                sqlComm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                    MessageBox.Show("Nickname apps exception " + apex.ErrorCode.ToString() + "  +++  \n" + apex.Data.ToString() + "  +++  \n" + apex.Message.ToString() + "  +++  \n" + apex.Reason.ToString() + "  +++  \n" + apex.Source.ToString());
                    if (apex.ErrorCode == "1301")
                    {
                        // this error about a non existent entiry seems to indicate the nickname is already created
                        complete = true;                        
                    }
                    i++;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Nickname issue " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
            string userNickName = "";
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
                        if (gusersyn.User_password_short_fix_checkbox == true && password.Length < 6)
                        {
                            password = GetPassword(14);
                        }



                        //Create a new user.
                        UserEntry insertedEntry = service.CreateUser(studentID, first_name, last_name, password);

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

                        log.transactions.Add("Added Gmail user " + studentID + "@" + gusersyn.Admin_domain + " Aliased as " + userNickName + "@" + gusersyn.Admin_domain);
                    }
                    catch (AppsException e)
                    {
                        log.errors.Add("Failed adding Gmail user " + studentID + "@" + gusersyn.Admin_domain + " Aliased as " + userNickName + "@" + gusersyn.Admin_domain + " failed " + e.Message.ToString() + " reason " + e.Reason.ToString());
                    }
                    catch (Exception ex)
                    {
                        log.errors.Add("Failed adding Gmail user " + studentID + "@" + gusersyn.Admin_domain + " Aliased as " + userNickName + "@" + gusersyn.Admin_domain + " failed " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                log.errors.Add("Issue adding gmail users datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                log.errors.Add("Failed Delete gmail account " + userID + " exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                        log.transactions.Add("Updated " + System.Web.HttpUtility.UrlEncode(usersToUpdate[gusersyn.User_StuID].ToString()).Replace("+", " ").Replace("*", "%2A") + " because of name change. New Name is " + gmailUser.Name.FamilyName.ToString() + ", " + gmailUser.Name.GivenName.ToString());
                        userNickName = GetNewUserNickname(service, gmailUser.Login.UserName, gmailUser.Name.GivenName, middlename, gmailUser.Name.FamilyName, 0, false);
                        log.transactions.Add("Added New Alias for " + gmailUser.Login.UserName + "@" + gusersyn.Admin_domain + " Aliased as " + userNickName + "@" + gusersyn.Admin_domain);
                    }
                    catch (Exception ex)
                    {
                        log.errors.Add("Failed update gmail account " + System.Web.HttpUtility.UrlEncode(usersToUpdate[gusersyn.User_StuID].ToString()).Replace("+", " ").Replace("*", "%2A").Replace("!", "%21").Replace("(", "%28").Replace(")", "%29").Replace("'", "%27").Replace("_", "%5f").Replace(" ", "%20").Replace("%", "_") + " exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
                    }

                }
            }
            catch (Exception ex)
            {
                log.errors.Add("Issue updating gmail users datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                        log.transactions.Add("Created send as alias " + (string)userNicknames[sendASFieldName] + " for userlogin " + (string)userNicknames[gusersyn.User_StuID]);
                    }
                    catch (Exception ex)
                    {
                        log.errors.Add("Failed user send as creation " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                log.errors.Add("Issue creating send as datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                log.errors.Add("failed to pull gmail user list exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                log.errors.Add("failed to pull gmail nickname list exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                log.errors.Add("failed to move ad object from " + objectLocation + " to " + newLocation + " exception " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
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
                log.errors.Add("failed to pull " + attributeName + " on object " + objectDn);
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
                log.errors.Add("Failure to create AD user account " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());

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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString() + "issue create user LDAP://CN=" + System.Web.HttpUtility.UrlEncode(properties["CN"].ToString()).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + "\n" + ex.StackTrace.ToString());
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
                                log.transactions.Add("Deleting user nickname " + nicknameEntry.Nickname.Name.ToString());
                            }
                        }
                        catch
                        {
                            log.errors.Add("Error deleting user nickname " + nicknames[userNicknameField].ToString());
                        }
                    }
                    // if the nickname is not found create the new nickname
                    if (foundnickname == false)
                    {
                        try
                        {
                            service.CreateNickname(nicknames[usernameFieldName].ToString(), nicknames[userNicknameField].ToString());
                            log.transactions.Add("Creating user nickname " + nicknames[userNicknameField].ToString() + " for user " + nicknames[usernameFieldName].ToString());
                        }
                        catch
                        {
                            log.errors.Add("Error adding user nickname " + nicknames[userNicknameField].ToString() + " for user " + nicknames[usernameFieldName].ToString());
                        }
                    }
                    // reset all variables
                    foundnickname = false;
                    i = 0;
                    nicknamecount = 0;
                }
            }
            catch(Exception ex)
            {
                log.errors.Add("Issue updating nicknames datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
            }
        }
    }


    //Logic code for execution of the synchronization
    public class ObjectADSqlsyncGroup
    {
        public void ExecuteGroupSync(GroupSynch groupsyn, ToolSet tools, LogFile log)
        {
            //string debug = "";
            //SqlDataReader debugreader;
            //ArrayList debuglist = new ArrayList();
            //int debugfieldcount;
            //string debugrecourdcount;
            //SqlCommand sqldebugComm;
            //int i;
            StopWatch time = new StopWatch();


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
            SqlConnection sqlConn = new SqlConnection("Data Source=" + groupsyn.DataServer + ";Initial Catalog=" + groupsyn.DBCatalog + ";Integrated Security=SSPI;Connect Timeout=360");



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


            try
            {
                sqlComm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
                throw;
            }


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
            groupsDataTable = tools.EnumerateGroupsInOUDataTable("OU=" + groupapp + "," + groupOU, adUpdateKeys, adGroupsTable, log);
            //time.Stop();

            //MessageBox.Show("got " + groupsLinkedList.Count + "groups from ou in " + time.GetElapsedTime());
            // insert groups from AD into a temp table
            if (groupsDataTable.Rows.Count > 0)
            {
                //time.Start();
                groupsTable = tools.Create_Table(groupsDataTable, adGroupsTable, sqlConn, log);
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
                add = tools.QueryNotExists(sqlgroupsTable, groupsTable, sqlConn, groupsyn.Group_CN, adUpdateKeys[1].ToString(), log);

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
                delete = tools.QueryNotExists(groupsTable, sqlgroupsTable, sqlConn, adUpdateKeys[1].ToString(), groupsyn.Group_CN, log);
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
                // make the list of fields for the sql to check when updating note these fields must be in the same order as the AD update keys
                sqlUpdateKeys.Add(groupsyn.Group_sAMAccount);
                sqlUpdateKeys.Add(groupsyn.Group_CN);
                //time.Start();
                // update assumes the both ADupdateKeys and SQLupdateKeys have the same fields, listed in the same order check  call to EnumerateGroupsInOU if this is wrong should be sAMAccountName, CN matching the SQL order
                update = tools.CheckUpdate(sqlgroupsTable, groupsTable, groupsyn.Group_CN, adUpdateKeys[1].ToString(), sqlUpdateKeys, adUpdateKeys, sqlConn, log);
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
                        groupDN = tools.GetObjectDistinguishedName(objectClass.group, returnType.distinguishedName, groupObject["CN"], dc, log);
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
            // we didn't find any records in AD so there is no need for the Update or delete logic to run
            else
            {
                sqlComm = new SqlCommand("SELECT * FROM " + sqlgroupsTable, sqlConn);
                try
                {
                    add = sqlComm.ExecuteReader();
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
                catch (Exception ex)
                {
                    log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
                }

                //time.Start();
                // i = 0;

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
            try
            {
                sqlComm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
            }

            // populate datatable with users from AD groups by looping thru the list of groups from SQL and loading the cross referenced AD group members
            sqlComm = new SqlCommand("SELECT " + groupsyn.Group_CN + " FROM " + sqlgroupsTable, sqlConn);
            try
            {
                sqlgroups = sqlComm.ExecuteReader();
                while (sqlgroups.Read())
                {
                    // hopefully merge acts as an append
                    ADusers.Merge(tools.EnumerateUsersInGroupDataTable((string)sqlgroups[0], ",OU=" + groupapp + "," + groupOU, groupsyn.User_sAMAccount, groupsyn.User_Group_Reference, ADgroupMembersTable, log));
                }
                sqlgroups.Close();
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
            }


            // make the temp table for ou comparisons
            tools.Create_Table(ADusers, ADgroupMembersTable, sqlConn, log);



            //debug = " total users in groups from SQL \n";
            //sqldebugComm = new SqlCommand("select top 20 * FROM " + sqlgroupMembersTable, sqlConn);
            //debugreader = sqldebugComm.ExecuteReader();
            //debugfieldcount = debugreader.FieldCount;
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
            //sqldebugComm = new SqlCommand("select count(" + ADusers.Columns[0].ColumnName + ") FROM " + sqlgroupMembersTable, sqlConn);
            //debugreader.Close();
            //debugrecourdcount = sqldebugComm.ExecuteScalar().ToString();
            //MessageBox.Show("table " + sqlgroupMembersTable + " has " + debugrecourdcount + " records \n " + debugfieldcount + " fields \n sample data" + debug);

            //debug = " total users in groups from AD \n";
            //sqldebugComm = new SqlCommand("select top 20 * FROM " + ADgroupMembersTable, sqlConn);
            //debugreader = sqldebugComm.ExecuteReader();
            //debugfieldcount = debugreader.FieldCount;
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
            //sqldebugComm = new SqlCommand("select count(" + ADusers.Columns[0].ColumnName + ") FROM " + ADgroupMembersTable, sqlConn);
            //debugreader.Close();
            //debugrecourdcount = sqldebugComm.ExecuteScalar().ToString();
            //MessageBox.Show("table " + ADgroupMembersTable + " has " + debugrecourdcount + " records \n " + debugfieldcount + " fields \n sample data" + debug);



            // compare and add/remove
            add = tools.QueryNotExists(sqlgroupMembersTable, ADgroupMembersTable, sqlConn, groupsyn.User_Group_Reference, ADusers.Columns[1].ColumnName, log);
            try
            {
            while (add.Read())
            {
                tools.AddUserToGroup((string)add[0], "CN=" + (string)add[1] + ",OU=" + groupapp + "," + groupOU, log);
                // log.transactions.Add("User added ;" + (string)add[0] + ",OU=" + groupapp + "," + groupOU + ";" + (string)add[1]);
                groupObject.Clear();
            }
            }
            catch (Exception ex)
            {
                log.errors.Add("Issue adding group datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
            }
            add.Close();
            SqlCommand sqlComm2 = new SqlCommand();
            string recordCount = "";
            sqlComm2 = new SqlCommand("select count(" + groupsyn.Group_sAMAccount + ") FROM " + sqlgroupMembersTable, sqlConn);
            recordCount =  sqlComm2.ExecuteScalar().ToString();
            sqlComm2.Dispose();

            if (recordCount != "0")
            {
                delete = tools.QueryNotExists(ADgroupMembersTable, sqlgroupMembersTable, sqlConn, ADusers.Columns[1].ColumnName, groupsyn.User_Group_Reference, log);
                // delete groups in AD
                while (delete.Read())
                {

                    tools.RemoveUserFromGroup((string)delete[0], (string)delete[1], log);
                    // log.transactions.Add("User removed ;" + (string)delete[adUpdateKeys[1].ToString()].ToString().Trim() + ",OU=" + groupapp + groupOU);

                }
                delete.Close();
            }
            sqlConn.Close();
        }
        public void ExecuteUserSync(UserSynch usersyn, ToolSet tools, LogFile log)
        {
            //string debug = "";
            //SqlDataReader debugReader;
            //int debugFieldCount;
            //string debugRecordCount;
            //SqlCommand sqlDebugComm;
            //int j;
            ArrayList debugList = new ArrayList();
            int i;
            StopWatch time = new StopWatch();


            string baseOU = usersyn.BaseUserOU;
            string DC = baseOU.Substring(baseOU.IndexOf("DC"));
            string sqlForCustomFields = "";

            SqlDataReader add;
            SqlDataReader delete;
            SqlDataReader update;

            ArrayList completeSqlKeys = new ArrayList();
            ArrayList completeADKeys = new ArrayList();
            ArrayList adUpdateKeys = new ArrayList();
            ArrayList sqlUpdateKeys = new ArrayList();
            ArrayList extraFieldsToReturn = new ArrayList();
            ArrayList fields = new ArrayList();
            Dictionary<string, string> userObject = new Dictionary<string, string>();
            SqlConnection sqlConn = new SqlConnection("Data Source=" + usersyn.DataServer + ";Initial Catalog=" + usersyn.DBCatalog + ";Integrated Security=SSPI;Connect Timeout=360");


            string sqlUsersTable = "#sqlusersTable";
            string adUsersTable = "#ADusersTable";
            //SqlDataReader sqlusers;
            SqlCommand sqlComm;
            SqlCommand sqlComm2;
            string recordCount = "";
            DataTable adUsers = new DataTable();



            sqlConn.Open();
            // creat initial ou's; will log a warning out if they already exist
            tools.CreateOURecursive(usersyn.BaseUserOU, log);
            tools.CreateOURecursive(usersyn.UserHoldingTank, log);

            // setup extentions for the user accounts to go in to the right ou's
            userObject.Add("sAMAccountName", usersyn.UniversalGroup.Remove(0, 3).Remove(usersyn.UniversalGroup.IndexOf(",") - 3));
            userObject.Add("CN", usersyn.UniversalGroup.Remove(0, 3).Remove(usersyn.UniversalGroup.IndexOf(",") - 3));
            userObject.Add("description", "Universal Group For Users");
            // creates the group if it does not exist
            tools.CreateGroup(usersyn.UniversalGroup.Remove(0, usersyn.UniversalGroup.IndexOf(",") + 1), userObject, log);


            // need to add this field first to use as a primary key when checking for existance in AD
            completeSqlKeys.Add("sAMAccountName");
            completeSqlKeys.Add("CN");
            completeSqlKeys.Add("sn");
            completeSqlKeys.Add("givenname");
            completeSqlKeys.Add("homephone");
            completeSqlKeys.Add("st");
            completeSqlKeys.Add("streetaddress");
            completeSqlKeys.Add("l");
            completeSqlKeys.Add("postalcode");
            // ?????? MIGHT NOT BE USED


            // Lets make the SQL fields to check for update
            sqlUpdateKeys.Add("sn");
            sqlUpdateKeys.Add("givenname");
            sqlUpdateKeys.Add("homephone");
            sqlUpdateKeys.Add("st");
            sqlUpdateKeys.Add("streetaddress");
            sqlUpdateKeys.Add("l");
            sqlUpdateKeys.Add("postalcode");



            // Lets make the Active Directory Keys as well
            completeADKeys.Add("CN");
            completeADKeys.Add("sn");
            completeADKeys.Add("givenname");
            completeADKeys.Add("homephone");
            completeADKeys.Add("st");
            completeADKeys.Add("streetaddress");
            completeADKeys.Add("l");
            completeADKeys.Add("postalcode");
            completeADKeys.Add("distinguishedName");

            // Lets make the Active Directory fields to check for update
            adUpdateKeys.Add("sn");
            adUpdateKeys.Add("givenname");
            adUpdateKeys.Add("homephone");
            adUpdateKeys.Add("st");
            adUpdateKeys.Add("streetaddress");
            adUpdateKeys.Add("l");
            adUpdateKeys.Add("postalcode");

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
            if (usersyn.User_where == "")
            {
                sqlComm = new SqlCommand("SELECT DISTINCT RTRIM(" + usersyn.User_sAMAccount + ") AS sAMAccountName" +
                    ", RTRIM(" + usersyn.User_sAMAccount + ") AS CN" +
                    ", RTRIM(" + usersyn.User_Lname + ") AS sn" +
                    ", RTRIM(" + usersyn.User_Fname + ") AS givenname" +
                    ", RTRIM(" + usersyn.User_Mobile + ") AS homephone" +
                    ", RTRIM(" + usersyn.User_State + ") AS st" +
                    ", RTRIM(" + usersyn.User_Address + ") AS streetaddress" +
                    //", RTRIM(" + usersyn.User_mail + ") AS mail" +
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
                    ", RTRIM(" + usersyn.User_State + ") AS st" +
                    ", RTRIM(" + usersyn.User_Address + ") AS streetaddress" +
                    //", RTRIM(" + usersyn.User_mail + ") AS mail" +
                    ", RTRIM(" + usersyn.User_city + ") AS l" +
                    ", RTRIM(" + usersyn.User_Zip + ") AS postalcode" +
                    ", RTRIM(" + usersyn.User_password + ") AS password" +
                    sqlForCustomFields +
                    " INTO " + sqlUsersTable + " FROM " + usersyn.User_dbTable +
                    " WHERE " + usersyn.User_where, sqlConn);
            }
            try
            {
                sqlComm.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
                throw;
            }

            // go grab all the users from AD
            adUsers = tools.EnumerateUsersInOUDataTable(usersyn.BaseUserOU, completeADKeys, adUsersTable, SearchScope.OneLevel, log);
            if (adUsers.Rows.Count > 0)
            {
                // make the temp table for ou comparisons
                tools.Create_Table(adUsers, adUsersTable, sqlConn, log);



                //************************************************************
                //                          START
                //                   DEBUG AND TEST DATA
                //
                //************************************************************
                //debug = " total users from sql \n";
                //sqlDebugComm = new SqlCommand("select top 20 * FROM " + sqlUsersTable, sqlConn);
                //debugReader = sqlDebugComm.ExecuteReader();
                //debugFieldCount = debugReader.FieldCount;
                //for (i = 0; i < debugFieldCount; i++)
                //{
                //    debug += debugReader.GetName(i) + ", ";
                //}
                //debug += "\n";
                //while (debugReader.Read())
                //{
                //    for (i = 0; i < debugFieldCount; i++)
                //    {
                //        debug += (string)debugReader[i].ToString() + ", ";
                //    }
                //    debug += "\n";
                //}
                //sqlDebugComm = new SqlCommand("select count(sAMAccountName) FROM " + sqlUsersTable, sqlConn);
                //debugReader.Close();
                //debugRecordCount = sqlDebugComm.ExecuteScalar().ToString();
                //MessageBox.Show("table " + sqlUsersTable + " has " + debugRecordCount + " records \n " + debugFieldCount + " fields \n sample data" + debug);


                //debug = "";
                //debug = " total users from AD \n";
                //sqlDebugComm = new SqlCommand("select top 20 * FROM " + adUsersTable, sqlConn);
                //debugReader = sqlDebugComm.ExecuteReader();
                //debugFieldCount = debugReader.FieldCount;
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
                //sqlDebugComm = new SqlCommand("select count(" + adUsers.Columns[0].ColumnName + ") FROM " + adUsersTable, sqlConn);
                //debugReader.Close();
                //debugRecordCount = sqlDebugComm.ExecuteScalar().ToString();
                //MessageBox.Show("table " + adUsersTable + " has " + debugRecordCount + " records \n " + debugFieldCount + " fields \n sample data" + debug);
                //************************************************************
                //                            END
                //                   DEBUG AND TEST DATA
                //
                //************************************************************






                // compare and add/remove
                add = tools.QueryNotExists(sqlUsersTable, adUsersTable, sqlConn, "sAMAccountName", adUsers.Columns[0].ColumnName, log);


                //debug = "Gunna Add stuff \n";
                //debugFieldCount = add.FieldCount;
                //for (i = 0; i < debugFieldCount; i++)
                //{
                //    debug += add.GetName(i) + ", ";
                //}
                //debug += "\n";
                //int j = 0;
                //while (add.Read() && j < 20)
                //{
                //    for (i = 0; i < debugFieldCount; i++)
                //    {
                //        debug += (string)add[i] + ", ";
                //    }
                //    debug += "\n";
                //    j++;
                //}

                ////debugReader.Close();
                //MessageBox.Show("table " + adUsersTable + "\n " + debugFieldCount + " fields \n sample data \n" + debug);

                tools.CreateUserAccount(usersyn.UserHoldingTank, add, usersyn.UniversalGroup, usersyn, log);
                add.Close();

      
                sqlComm2 = new SqlCommand("select count(sAMAccountName) FROM " + sqlUsersTable, sqlConn);
                recordCount =  sqlComm2.ExecuteScalar().ToString();
                sqlComm2.Dispose();

                if (recordCount != "0")
                {
                    delete = tools.QueryNotExists(adUsersTable, sqlUsersTable, sqlConn, usersyn.User_sAMAccount, completeADKeys[0].ToString(), log);

                    //debug = "Gunna delete stuff \n field names are \n";
                    //debugFieldCount = delete.FieldCount;
                    //for (i = 0; i < debugFieldCount; i++)
                    //{
                    //    debug += delete.GetName(i) + ", ";
                    //}
                    //debug += "\n data \n";
                    ////int j = 0;
                    //j = 0;
                    //while (delete.Read() && j < 20)
                    //{
                    //    for (i = 0; i < debugFieldCount; i++)
                    //    {
                    //        debug += (string)delete[i] + ", ";
                    //    }
                    //    debug += "\n";
                    //    j++;
                    //}

                    //MessageBox.Show(debug);

                    // delete users in AD
                    try
                    {
                        while (delete.Read())
                        {

                            tools.DeleteUserAccount((string)delete["distinguishedname"], log);
                            // log.transactions.Add("User removed ;" + (string)delete[adUpdateKeys[1].ToString()].ToString().Trim()); 
                        }
                    }
                    catch (Exception ex)
                    {
                        log.errors.Add("Issue deleting AD users datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
                    }
                    delete.Close();
                }


                // add the extra fields in form ".field ,"
                extraFieldsToReturn.Add(adUsersTable + ".distinguishedname ,");

                update = tools.CheckUpdate(sqlUsersTable, adUsersTable, usersyn.User_sAMAccount, "CN", sqlUpdateKeys, adUpdateKeys, extraFieldsToReturn, sqlConn, log);

                //debug = "Gunna update stuff \n field names are \n";
                //debugFieldCount = update.FieldCount;
                //for (i = 0; i < debugFieldCount; i++)
                //{
                //    debug += update.GetName(i) + ", ";
                //}
                //debug += "\n data \n";
                ////int j = 0;
                //j = 0;
                //while (update.Read() && j < 20)
                //{
                //    for (i = 0; i < debugFieldCount; i++)
                //    {
                //        debug += (string)update[i] + ", ";
                //    }
                //    debug += "\n";
                //    j++;
                //}

                //MessageBox.Show(debug);

                tools.UpdateUsers(update, log);
                // update users in ad
                // last record which matches the primary key is the one which gets inserted into the database
                //while (update.Read())
                //{
                //    // any duplicate records will attempt to be updated if slow runtimes are a problem this might be an issue
                //    if (tools.Exists((string)update["distinguishedName"].ToString().Trim()) == true)
                //    {
                //        // user exists in place just needs updating
                //        tools.UpdateGroup("OU=" + groupapp + "," + groupOU, groupObject, log);
                //    }
                //    else
                //    {
                //        // find it its on the server somewhere we will log the exception
                //        userDN = tools.GetObjectDistinguishedName(objectClass.user, returnType.distinguishedName, (string)update["sAMAccountName"].ToString().Trim(), dc);
                //        log.errors.Add("User could not be updated user probabally should be in ; " + "OU=" + groupapp + "," + groupOU + " ; but was found in ; " + userDN);
                //    }

                //}
                update.Close();
            }
            // did not find any records in AD
            else
            {
                // add the users without doing additional checks
                tools.Create_Table(adUsers, adUsersTable, sqlConn, log);
                add = tools.QueryNotExists(sqlUsersTable, adUsersTable, sqlConn, "sAMAccountName", adUsers.Columns[0].ColumnName, log);
                tools.CreateUserAccount(usersyn.UserHoldingTank, add, usersyn.UniversalGroup, usersyn, log);
                add.Close();
            }
            sqlConn.Close();
        }
    }
    public class ObjectADGoogleSync
    {
        public void EmailUsersSync(GmailUsers gusersyn, ToolSet tools, LogFile log)
		{
            // MessageBox.Show("gmail " + gusersyn.Admin_password + " " + gusersyn.Admin_domain + " " + gusersyn.Admin_user + " " + gusersyn.DataServer + " " + gusersyn.DBCatalog + " " + gusersyn.User_ad_OU + " " + gusersyn.User_Datasource + " " + gusersyn.User_dbTable + " " + gusersyn.User_Fname + " " + gusersyn.User_Lname + " " + gusersyn.User_Mname + " " + gusersyn.User_password + " " + gusersyn.User_password_short_fix_checkbox.ToString() + " " + gusersyn.User_password_generate_checkbox.ToString() + " " + gusersyn.User_StuID + " " + gusersyn.User_table_view + " " + gusersyn.User_where + " " + gusersyn.Writeback_AD_checkbox.ToString() + " " + gusersyn.Writeback_ad_OU + " " + gusersyn.Writeback_DB_checkbox.ToString() + " " + gusersyn.Writeback_email_field + " " + gusersyn.Writeback_primary_key + " " + gusersyn.Writeback_secondary_email_field + " " + gusersyn.Writeback_table + " " + gusersyn.Writeback_transfer_email_checkbox.ToString() + " " + gusersyn.Writeback_where_clause);
			// Email addresses are static so only the names can be updated. passwords will be ignored
			// appservice variables will come from a config designed ot hold its data (sql and Gmail login)
            int i = 0;
            string userDN = "";
			AppsService service = new AppsService(gusersyn.Admin_domain, gusersyn.Admin_user + "@" + gusersyn.Admin_domain, gusersyn.Admin_password);
			ArrayList completeSqlKeys = new ArrayList();
			ArrayList completeGmailKeys = new ArrayList();
			ArrayList gmailUpdateKeys = new ArrayList();
			ArrayList sqlUpdateKeys = new ArrayList();
            ArrayList adUpdateKeys = new ArrayList();
            ArrayList additionalKeys = new ArrayList();
                                                             

			SqlDataReader add;
			//SqlDataReader delete;
			SqlDataReader update;
            SqlConnection sqlConn = new SqlConnection("Data Source=" + gusersyn.DataServer + ";Initial Catalog=" + gusersyn.DBCatalog + ";Integrated Security=SSPI;Connect Timeout=360;");


            //string sqlUsersTable = "#sqlusersTable";
            //string gmailUsersTable = "#gmailusersTable";
            string sqlUsersTable = "FHC_TEST_sqlusersTable";
            string gmailUsersTable = "FHC_TEST_gmailusersTable";
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
            tools.DropTable(sqlUsersTable, sqlConn, log);
            tools.DropTable(gmailUsersTable, sqlConn, log);

            // this statement picks the datasource SQL vs AD and sets up the temp table
            if (gusersyn.User_Datasource == "database")
            {
                // grab users data from sql
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
                    sqlComm.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
                    // If we add deletion we need to fail out if this is blank
                    //throw;
                }
            }
            else
            {
                adUsers = tools.EnumerateUsersInOUDataTable(gusersyn.User_ad_OU, completeSqlKeys, sqlUsersTable, SearchScope.OneLevel, log);
                // build the database temp table from the users retrieved into adUsers
                tools.Create_Table(adUsers, sqlUsersTable, sqlConn, log);
            }   

            // go grab all the users from Gmail from the database
            gmailUsers = tools.Get_Gmail_Users(service, gusersyn, gmailUsersTable, log);
            // make the temp table for ou comparisons
            tools.Create_Table(gmailUsers, gmailUsersTable, sqlConn, log);

            

			// compare and add/remove
			add = tools.QueryNotExists(sqlUsersTable, gmailUsersTable, sqlConn, gusersyn.User_StuID, gmailUsers.Columns[0].ColumnName, log);

            //SqlDataReader debugReader;
            //string debug = "";
            //int debugFieldCount = 0;
            
            //debug = "Gunna Add stuff \n";
            //debugFieldCount = add.FieldCount;
            //for (i = 0; i < debugFieldCount; i++)
            //{
            //    debug += add.GetName(i) + ", ";
            //}
            //debug += "\n";
            //int j = 0;
            //while (add.Read() && j < 20)
            //{
            //    for (i = 0; i < debugFieldCount; i++)
            //    {
            //        debug += (string)add[i] + ", ";
            //    }
            //    debug += "\n";
            //    j++;
            //}

            // debugReader.Close();
            // MessageBox.Show("Gmail users to add \n " + debugFieldCount + " fields \n sample data" + debug);

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

            update = tools.CheckUpdate(sqlUsersTable, gmailUsersTable, gusersyn.User_StuID, gmailUsers.Columns[0].ColumnName, sqlUpdateKeys, gmailUpdateKeys, additionalKeys, sqlConn, log);


            //SqlDataReader debugReader;
            //string debug = "";
            //int debugFieldCount = 0;

            //debug = "Gunna update stuff \n";
            //debugFieldCount = update.FieldCount;
            //for (i = 0; i < debugFieldCount; i++)
            //{
            //    debug += update.GetName(i) + ", ";
            //}
            //debug += "\n";
            //int j = 0;
            //try
            //{
            //    while (update.Read() && j < 20)
            //    {
            //        for (i = 0; i < debugFieldCount; i++)
            //        {
            //            debug += (string)update[i] + ", ";
            //        }
            //        debug += "\n";
            //        j++;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    log.warnings.Add("Issue creating debug data " + ex.Message.ToString());
            //}

            ////debugReader.Close();
            //MessageBox.Show("Gmail users to update \n " + debugFieldCount + " fields \n sample data" + debug);


            tools.UpdateGmailUser(service, gusersyn, update, log);
            update.Close();
            




            // ***********************************
            // ** Start writeback features
            // ***********************************


            //string nicknamesFromGmailTable = "#gmailNicknamesTable";
            //string loginWithoutNicknamesTable = "#loginsWONicknamesTable";
            //string adNicknamesTable = "#adNicknamesTable";
            //string sqlNicknamesTable = "#sqlNicknamesTable";
            //string nicknamesToUpdateDBTable = "#nicknamesToUpdateDB";
            string nicknamesFromGmailTable = "FHC_TEST_gmailNicknamesTable";
            string loginWithoutNicknamesTable = "FHC_TEST_loginsWONicknamesTable";
            string adNicknamesTable = "FHC_TEST_adNicknamesTable";
            string sqlNicknamesTable = "FHC_TEST_sqlNicknamesTable";
            string nicknamesToUpdateDBTable = "FHC_TEST_nicknamesToUpdateDB";
            string nicknamesFilteredForDuplicatesTable = "FHC_TEST_nicknamesFilteredDuplicates";
            string dc = gusersyn.Writeback_ad_OU.Substring(gusersyn.Writeback_ad_OU.IndexOf("DC"));
            string userNickName = "";
            SqlDataReader nicknamesToAddToDatabase;
            SqlDataReader nicknamesToAddToAD;
            SqlDataReader lostNicknames;
            ArrayList nicknameKeys = new ArrayList();
            ArrayList sqlkeys = new ArrayList();
            ArrayList adPullKeys = new ArrayList();
            ArrayList adMailUpdateKeys = new ArrayList();
            ArrayList nicknameKeysAndTable = new ArrayList();
            ArrayList sqlkeysAndTable = new ArrayList();
            ArrayList keywordFields = new ArrayList(); //fields fror checking against nickname to see how close it is to the real data
            ArrayList selectFields = new ArrayList(); //fields from nicknamesFromGmailTable to bring back
            DataTable nicknames = new DataTable();
            SqlDataReader sendAsAliases;


            // housecleaning
            tools.DropTable(nicknamesFromGmailTable, sqlConn, log);
            tools.DropTable(loginWithoutNicknamesTable, sqlConn, log);
            tools.DropTable(adNicknamesTable, sqlConn, log);
            tools.DropTable(sqlNicknamesTable, sqlConn, log);
            tools.DropTable(nicknamesToUpdateDBTable, sqlConn, log);
            tools.DropTable(nicknamesFilteredForDuplicatesTable, sqlConn, log);

            if (gusersyn.Writeback_AD_checkbox == true || gusersyn.Writeback_DB_checkbox == true)
            {
                // DATABASE writeback

                // Make preperations to pull all data into seperate tables
                // build sql to run to get gmail user nicknames
                // execute data pull for SQL nicknames used in writeback to SQL database
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
                        sqlComm.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        log.errors.Add("Failed SQL command " + sqlComm.CommandText.ToString() + " error " + ex + "\n" + ex.StackTrace.ToString());
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
                adUsers.Clear();
                adUsers = tools.EnumerateUsersInOUDataTable(gusersyn.User_ad_OU, adPullKeys, adNicknamesTable, SearchScope.OneLevel, log);
                tools.Create_Table(adUsers, adNicknamesTable, sqlConn, log);


                // get list of nicknames from gmail
                nicknames.Clear();
                nicknames = tools.Get_Gmail_Nicknames(service, gusersyn, nicknamesFromGmailTable, log);
                tools.Create_Table(nicknames, nicknamesFromGmailTable, sqlConn, log);

                // get list of users from gmail this may have changed when we ran the update
                gmailUsers.Clear();
                tools.DropTable(gmailUsersTable, sqlConn, log);
                gmailUsers = tools.Get_Gmail_Users(service, gusersyn, gmailUsersTable, log);
                tools.Create_Table(gmailUsers, gmailUsersTable, sqlConn, log);


                // check which nicknames do not have a an associated user with them 
                // cross reference for null userID's in nickname service.RetrieveAllNicknames table with list of all userlogin userID's from gmail service.RetrieveAllUsers
                tools.QueryNotExistsIntoNewTable(gmailUsersTable, nicknamesFromGmailTable, loginWithoutNicknamesTable, sqlConn, gusersyn.User_StuID, nicknames.Columns[0].ColumnName, log);


                //************************************************************
                //                          START
                //                   DEBUG AND TEST DATA
                //
                //************************************************************


                //SqlDataReader debugReader;
                //string debug = "";
                //SqlCommand sqlDebugComm;
                //string firstfield = "";
                //string debugRecordCount = "";
                //int debugFieldCount = 0;

               // try
               // {

               //     //string sqlUsersTable = "#sqlusersTable";
               //     debug = " Users from sql \n";
               //     sqlDebugComm = new SqlCommand("select top 40 * FROM " + sqlUsersTable, sqlConn);
               //     debugReader = sqlDebugComm.ExecuteReader();
               //     debugFieldCount = debugReader.FieldCount;
               //     firstfield = debugReader.GetName(0);
               //     for (i = 0; i < debugFieldCount; i++)
               //     {
               //         debug += debugReader.GetName(i) + ", ";
               //     }
               //     debug += "\n";
               //     while (debugReader.Read())
               //     {
               //         for (i = 0; i < debugFieldCount; i++)
               //         {
               //             debug += (string)debugReader[i].ToString() + ", ";
               //         }
               //         debug += "\n";
               //     }
               //     sqlDebugComm = new SqlCommand("select count(" + firstfield + ") FROM " + sqlUsersTable, sqlConn);
               //     debugReader.Close();
               //     debugRecordCount = sqlDebugComm.ExecuteScalar().ToString();
               //     MessageBox.Show("table " + sqlUsersTable + " has " + debugRecordCount + " records \n " + debugFieldCount + " fields \n sample data" + debug);



               //     //string gmailUsersTable = "#gmailusersTable";
               //     debug = "";
               //     debug = " Users from gmail \n";
               //     sqlDebugComm = new SqlCommand("select top 40 * FROM " + gmailUsersTable, sqlConn);
               //     debugReader = sqlDebugComm.ExecuteReader();
               //     debugFieldCount = debugReader.FieldCount;
               //     firstfield = debugReader.GetName(0);
               //     for (i = 0; i < debugFieldCount; i++)
               //     {
               //         debug += debugReader.GetName(i) + ", ";
               //     }
               //     debug += "\n";
               //     while (debugReader.Read())
               //     {
               //         for (i = 0; i < debugFieldCount; i++)
               //         {
               //             debug += (string)debugReader[i] + ", ";
               //         }
               //         debug += "\n";
               //     }
               //     sqlDebugComm = new SqlCommand("select count(" + firstfield + ") FROM " + gmailUsersTable, sqlConn);
               //     debugReader.Close();
               //     debugRecordCount = sqlDebugComm.ExecuteScalar().ToString();
               //     MessageBox.Show("table " + gmailUsersTable + " has " + debugRecordCount + " records \n " + debugFieldCount + " fields \n sample data" + debug);


                ////string nicknamesFromGmailTable = "#gmailNicknamesTable";
                //debug = "";
                //debug = " Nicknames from gmail \n";
                //sqlDebugComm = new SqlCommand("select top 40 * FROM " + nicknamesFromGmailTable, sqlConn);
                //debugReader = sqlDebugComm.ExecuteReader();
                //debugFieldCount = debugReader.FieldCount;
                //firstfield = debugReader.GetName(0);
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
                //sqlDebugComm = new SqlCommand("select count(" + firstfield + ") FROM " + nicknamesFromGmailTable, sqlConn);
                //debugReader.Close();
                //debugRecordCount = sqlDebugComm.ExecuteScalar().ToString();
                //MessageBox.Show("table " + nicknamesFromGmailTable + " has " + debugRecordCount + " records \n " + debugFieldCount + " fields \n sample data" + debug);


               //     //string loginWithoutNicknamesTable = "#loginsWONicknamesTable";
               //     debug = "";
               //     debug = " Logons without nicknames \n";
               //     sqlDebugComm = new SqlCommand("select top 40 * FROM " + loginWithoutNicknamesTable, sqlConn);
               //     debugReader = sqlDebugComm.ExecuteReader();
               //     debugFieldCount = debugReader.FieldCount;
               //     firstfield = debugReader.GetName(0);
               //     for (i = 0; i < debugFieldCount; i++)
               //     {
               //         debug += debugReader.GetName(i) + ", ";
               //     }
               //     debug += "\n";
               //     while (debugReader.Read())
               //     {
               //         for (i = 0; i < debugFieldCount; i++)
               //         {
               //             debug += (string)debugReader[i] + ", ";
               //         }
               //         debug += "\n";
               //     }
               //     sqlDebugComm = new SqlCommand("select count(" + firstfield + ") FROM " + loginWithoutNicknamesTable, sqlConn);
               //     debugReader.Close();
               //     debugRecordCount = sqlDebugComm.ExecuteScalar().ToString();
               //     MessageBox.Show("table " + loginWithoutNicknamesTable + " has " + debugRecordCount + " records \n " + debugFieldCount + " fields \n sample data" + debug);

               //     if (gusersyn.Writeback_AD_checkbox == true || gusersyn.Writeback_DB_checkbox == true)
               //     {
               //         if (gusersyn.Writeback_AD_checkbox == true)
               //         {
               //             //string adNicknamesTable = "#adNicknamesTable";
               //             debug = "";
               //             debug = " Nickname list from AD for writeback \n";
               //             sqlDebugComm = new SqlCommand("select top 40 * FROM " + adNicknamesTable, sqlConn);
               //             debugReader = sqlDebugComm.ExecuteReader();
               //             debugFieldCount = debugReader.FieldCount;
               //             firstfield = debugReader.GetName(0);
               //             for (i = 0; i < debugFieldCount; i++)
               //             {
               //                 debug += debugReader.GetName(i) + ", ";
               //             }
               //             debug += "\n";
               //             while (debugReader.Read())
               //             {
               //                 for (i = 0; i < debugFieldCount; i++)
               //                 {
               //                     debug += (string)debugReader[i] + ", ";
               //                 }
               //                 debug += "\n";
               //             }
               //             sqlDebugComm = new SqlCommand("select count(" + firstfield + ") FROM " + adNicknamesTable, sqlConn);
               //             debugReader.Close();
               //             debugRecordCount = sqlDebugComm.ExecuteScalar().ToString();
               //             MessageBox.Show("table " + adNicknamesTable + " has " + debugRecordCount + " records \n " + debugFieldCount + " fields \n sample data" + debug);
               //         }

               //         if (gusersyn.Writeback_DB_checkbox == true)
               //         {
                //string sqlNicknamesTable = "#sqlNicknamesTable";
                //debug = "";
                //debug = " Nicknames from sql fro writeback \n";
                //sqlDebugComm = new SqlCommand("select top 40 * FROM " + sqlNicknamesTable, sqlConn);
                //debugReader = sqlDebugComm.ExecuteReader();
                //debugFieldCount = debugReader.FieldCount;
                //firstfield = debugReader.GetName(0);
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
                //sqlDebugComm = new SqlCommand("select count(" + firstfield + ") FROM " + sqlNicknamesTable, sqlConn);
                //debugReader.Close();
                //debugRecordCount = sqlDebugComm.ExecuteScalar().ToString();
                //MessageBox.Show("table " + sqlNicknamesTable + " has " + debugRecordCount + " records \n " + debugFieldCount + " fields \n sample data" + debug);
               //         }
               //     }
               // }
               // catch (Exception ex)
               // {
               //     log.warnings.Add("Issue creating debug data some sort of sql error " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
               // }

                
                //************************************************************
                //                            END
                //                   DEBUG AND TEST DATA
                //
                //************************************************************







                // use retrieved list of users without nicknames and check for updates against list of users in main datasource
                // use the datatable from the view/table as the primary data source
                // this table is generated above during the user account addition and update section
                // sqlUsersTable will have AD or SQL no matter which we checked, there was only on variable used
                // Pulls back the account infromation from the add/update section for users without a nickname
                // loginWithoutNicknamesTable inherits key from from #gmailuserstable during QueryNotExistsIntoNewTable due to only pulling data from the gmailuserstable

                lostNicknames = tools.QueryInnerJoin(sqlUsersTable, loginWithoutNicknamesTable, gusersyn.User_StuID, gusersyn.User_StuID, sqlConn, log);
                // iterate lostnicknames and create nicknames
                try
                {
                    while (lostNicknames.Read())
                    {
                        userNickName = tools.GetNewUserNickname(service, lostNicknames[gusersyn.User_StuID.ToString()].ToString(), lostNicknames[gusersyn.User_Fname.ToString()].ToString(), lostNicknames[gusersyn.User_Mname.ToString()].ToString(), lostNicknames[gusersyn.User_Lname.ToString()].ToString(), 0, false);
                        log.transactions.Add("Added Gmail user " + lostNicknames[gusersyn.User_StuID.ToString()].ToString() + "@" + gusersyn.Admin_domain + " Aliased as " + userNickName + "@" + gusersyn.Admin_domain);
                    }
                }
                catch (Exception ex)
                {
                    log.errors.Add("Issue creating new nicknames datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
                }

                lostNicknames.Close();
                // get list of nicknames from gmail it may have changed in the last update. Need to clear teh temp table or we will get an error
                tools.DropTable(nicknamesFromGmailTable, sqlConn, log);
                nicknames = tools.Get_Gmail_Nicknames(service, gusersyn, nicknamesFromGmailTable, log);
                tools.Create_Table(nicknames, nicknamesFromGmailTable, sqlConn, log);


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


                    // Filter nicknames for duplicates into new table 
                    tools.SelectNicknamesClosestToActualNameIntoNewTable(nicknamesFromGmailTable, sqlUsersTable, nicknames.Columns[0].ColumnName, gusersyn.User_StuID, nicknamesFilteredForDuplicatesTable, selectFields, nicknames.Columns[1].ColumnName, keywordFields, sqlConn, log);
                    // Check filtered nicknames against the sql data to see which emails need updating and put into a table for the next operation
                    tools.CheckEmailUpdateIntoNewTable(nicknames.Columns[2].ColumnName, gusersyn.Writeback_email_field, nicknamesFilteredForDuplicatesTable, sqlNicknamesTable, nicknames.Columns[0].ColumnName, gusersyn.Writeback_primary_key, nicknamesToUpdateDBTable, nicknameKeys, sqlkeys, sqlConn, log);



                     //SqlDataReader debugReader;
                     //string debug = "";
                     //SqlCommand sqlDebugComm;
                     //string firstfield = "";
                     //string debugRecordCount = "";
                     //int debugFieldCount = 0;
                     //string sqlUsersTable = "#sqlusersTable";
                     //debug = " Users from sql \n";
                     //sqlDebugComm = new SqlCommand("select top 40 * FROM " + nicknamesToUpdateDBTable, sqlConn);
                     //debugReader = sqlDebugComm.ExecuteReader();
                     //debugFieldCount = debugReader.FieldCount;
                     //firstfield = debugReader.GetName(0);
                     //for (i = 0; i < debugFieldCount; i++)
                     //{
                     //    debug += debugReader.GetName(i) + ", ";
                     //}
                     //debug += "\n";
                     //while (debugReader.Read())
                     //{
                     //    for (i = 0; i < debugFieldCount; i++)
                     //    {
                     //        debug += (string)debugReader[i].ToString() + ", ";
                     //    }
                     //    debug += "\n";
                     //}
                     //sqlDebugComm = new SqlCommand("select count(" + firstfield + ") FROM " + nicknamesToUpdateDBTable, sqlConn);
                     //debugReader.Close();
                     //debugRecordCount = sqlDebugComm.ExecuteScalar().ToString();
                     //MessageBox.Show("table " + sqlUsersTable + " has " + debugRecordCount + " records \n " + debugFieldCount + " fields \n sample data" + debug);


                    // nicknamesToAddToDatabase = tools.QueryNotExists(nicknamesFromGmailTable, sqlNicknamesTable, sqlConn, nicknames.Columns[0].ColumnName, gusersyn.Writeback_primary_key, log);
                    // update email fields in database where we did not have an primary key in the second table
                    //tools.Mass_update_email_field(nicknamesToAddToDatabase, sqlConn, gusersyn, log);

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
                        tools.Mass_Email_Shift(gusersyn, nicknamesToUpdateDBTable, gusersyn.Writeback_table, nicknames.Columns[0].ColumnName, gusersyn.Writeback_primary_key, nicknameKeysAndTable, sqlkeysAndTable, gusersyn.Writeback_where_clause, sqlConn, log); 
                    }
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
                    nicknamesToAddToAD = tools.CheckUpdate(nicknamesFilteredForDuplicatesTable, adNicknamesTable, nicknames.Columns[0].ColumnName, adPullKeys[0].ToString(), nicknameKeys, adMailUpdateKeys, sqlConn, log);
                    //nicknamesToAddToAD = tools.QueryNotExists(nicknamesFromGmailTable, adNicknamesTable, sqlConn, nicknames.Columns[0].ColumnName, adPullKeys[0].ToString(), log);
                    // interate and update mail fields

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
                        log.errors.Add("Issue writing nickname to AD datareader is null " + ex.Message.ToString() + "\n" + ex.StackTrace.ToString());
                    }
                    nicknamesToAddToAD.Close();
                }
            

                // create send as aliases need to get data for users 

                additionalKeys.Clear();
                additionalKeys.Add(nicknamesFilteredForDuplicatesTable + "." + nicknames.Columns[2].ColumnName + ", ");
                sendAsAliases = tools.QueryInnerJoin(sqlUsersTable, nicknamesFilteredForDuplicatesTable, gusersyn.User_StuID, nicknames.Columns[0].ColumnName, additionalKeys, sqlConn, log);
                tools.CreateSendAs(gusersyn, sendAsAliases, nicknames.Columns[2].ColumnName, nicknames.Columns[2].ColumnName, log);
                

                //tools.DropTable(nicknamesFromGmailTable, sqlConn, log);
                //tools.DropTable(loginWithoutNicknamesTable, sqlConn, log);
                //tools.DropTable(adNicknamesTable, sqlConn, log);
                //tools.DropTable(sqlNicknamesTable, sqlConn, log);
                //tools.DropTable(nicknamesToUpdateDBTable, sqlConn, log);
            }                      
            //tools.DropTable(sqlUsersTable, sqlConn, log);
            //tools.DropTable(gmailUsersTable, sqlConn, log);
            //nicknamesToAddToDatabase.Close();
            //nicknamesToAddToAD.Close();
			sqlConn.Close();
		}
    }
}

