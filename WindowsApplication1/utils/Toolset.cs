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
using WindowsApplication1;
using WindowsApplication1.utils;

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
	public class GmailUsers
	{
		private String configUser_table_view;
		private String configUser_dbTable;
		private String configUser_where;
		private String configDataServer;
		private String configDBCatalog;
		private String configUser_Lname;
		private String configUser_Fname;
		private String configUser_MiddleName;
		private String configUser_StuID;
		private String configUser_password;
		private String configUser_adminUser;
		private String configUser_adminDomain;
		private String configUser_adminPassword;
		public GmailUsers()
		{
			configUser_table_view = "";
			configUser_dbTable = "";
			configUser_where = "";
			configDataServer = "";
			configDBCatalog = "";
			configUser_Lname = "";
			configUser_Fname = "";
			configUser_MiddleName = "";
			configUser_StuID = "";
			configUser_password = "";
			configUser_adminUser = "";
			configUser_adminDomain = "";
			configUser_adminPassword = "";
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

		public void load(Dictionary<string, string> dictionary)
		{
			dictionary.TryGetValue("configUser_table_view", out configUser_table_view);
			dictionary.TryGetValue("configUser_dbTable", out configUser_dbTable);
			dictionary.TryGetValue("configUser_where", out configUser_where);
			dictionary.TryGetValue("configDataServer", out configDataServer);
			dictionary.TryGetValue("configDBCatalog", out configDBCatalog);
			dictionary.TryGetValue("configUser_Lname", out configUser_Lname);
			dictionary.TryGetValue("configUser_Fname", out configUser_Fname);
			dictionary.TryGetValue("configUser_MiddleName", out configUser_MiddleName);
			dictionary.TryGetValue("configUser_StuID", out configUser_StuID);
			dictionary.TryGetValue("configUser_password", out configUser_password);
			dictionary.TryGetValue("configUser_adminUser", out configUser_adminUser);
			dictionary.TryGetValue("configUser_adminDomain", out configUser_adminDomain);
			dictionary.TryGetValue("configUser_adminPassword", out configUser_adminPassword);
		}

		public Dictionary<string, string> ToDictionary()
		{
			Dictionary<string, string> returnvalue = new Dictionary<string, string>();
			returnvalue.Add("configUser_table_view", configUser_table_view);
			returnvalue.Add("configUser_dbTable", configUser_dbTable);
			returnvalue.Add("configUser_where", configUser_where);
			returnvalue.Add("configDataServer", configDataServer);
			returnvalue.Add("configDBCatalog", configDBCatalog);
			returnvalue.Add("configUser_Lname", configUser_Lname);
			returnvalue.Add("configUser_Fname", configUser_Fname);
			returnvalue.Add("configUser_MiddleName", configUser_MiddleName);
			returnvalue.Add("configUser_StuID", configUser_StuID);
			returnvalue.Add("configUser_password", configUser_password);
			returnvalue.Add("configUser_adminUser", configUser_adminUser);
			returnvalue.Add("configUser_adminDomain", configUser_adminDomain);
			returnvalue.Add("configUser_adminPassword", configUser_adminPassword);
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
                catch
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
                    log.errors.Add(e.Message.ToString() + " error deleting LDAP://CN=" + name + "," + ouPath);
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
            // oupath holds the path for the AD OU to hold the Users 
            // users is a sqldatareader witht the required fields in it ("CN") other Datastructures would be easy to substitute 
            // groupDN is a base group which all new users get automatically inserted into

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
                catch (Exception e)
                {
                    string debugdata = "";
                    for (i = 0; i < fieldcount; i++)
                    {

                        debugdata += users.GetName(i) + "=" + System.Web.HttpUtility.UrlEncode((string)users[i]).Replace("+", " ").Replace("*", "%2A") + ", ";

                    }
                    log.errors.Add("issue create user LDAP://CN=" + System.Web.HttpUtility.UrlEncode((string)users["CN"]).Replace("+", " ").Replace("*", "%2A") + "," + ouPath + "\n" + debugdata + " failed field maybe " + name + " | " + e.Message.ToString());
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
        public void UpdateUsers(SqlDataReader users, LogFile log)
        {
            int fieldcount = 0;
            int i = 0;
            string name = "";
            fieldcount = users.FieldCount;
            while (users.Read())
            {
                try
                {
                    DirectoryEntry user = new DirectoryEntry(System.Web.HttpUtility.UrlEncode((string)users["distinguishedname"]).Replace("+", " ").Replace("*", "%2A"));
                    for (i = 0; i < fieldcount; i++)
                    {
                        name = users.GetName(i);
                        if (name != "password" && name != "CN")
                        {
                            if ((string)users[i] != "")
                            {
                                user.Properties[users.GetName(i)].Value = System.Web.HttpUtility.UrlEncode((string)users[i]).Replace("+", " ").Replace("*", "%2A");
                            }
                        }
                    }
                    user.CommitChanges();
                    log.transactions.Add("User updated |" + (string)users["CN"] + " ");
                }
                catch (Exception e)
                {
                    log.errors.Add("issue updating user " + System.Web.HttpUtility.UrlEncode((string)users["distinguishedname"]).Replace("+", " ").Replace("*", "%2A") + "\n" + e.Message.ToString());
                }

            }
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
        public SqlDataReader CheckUpdate(string table1, string table2, string pkey1, string pkey2, ArrayList compareFields1, ArrayList compareFields2, ArrayList additionalFields, SqlConnection sqlConn)
        {
            // additionalFields takes the field names " table.field,"
            // Assumes table1 holds the correct data and returns a data reader with the update fields columns from table1
            // compare fields 1 & 2 should have the same number of items or it is likely that all row will found needing updating
            // returns fields from comparefields
            // returns the rows which table2 differs from table1
            string compare1 = "";
            string compare2 = "";
            string fields = "";
            string additionalfields = "";
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
            foreach (string key in additionalFields)
            {
                additionalfields += key;
            }
            // remove trailing comma and + 
            compare2 = compare2.Remove(compare2.Length - 2);
            compare1 = compare1.Remove(compare1.Length - 2);
            fields = fields.Remove(fields.Length - 2);
            additionalfields = additionalfields.Remove(additionalfields.Length - 2);
            SqlCommand sqlComm;
            if (additionalFields.Count > 0)
            {
                sqlComm = new SqlCommand("SELECT " + fields + ", " + additionalfields + " FROM " + table1 + " INNER JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2 + " WHERE (" + compare2 + ") <> (" + compare1 + ")", sqlConn);
            }
            else
            {
                sqlComm = new SqlCommand("SELECT " + fields + " FROM " + table1 + " INNER JOIN " + table2 + " ON " + table1 + "." + pkey1 + " = " + table2 + "." + pkey2 + " WHERE (" + compare2 + ") <> (" + compare1 + ")", sqlConn);
            }
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
            // returns a list of all the AD fields in the schema for a user
            // modification to any active directory object type would be simple by changing the directory entry
            // NOTE: One place where managed ADSI (System.DirectoryServices) falls short is finding schema 
            // information from LDAP/AD objects. Finding information like mandatory and optional
            // properties simply cannot be done with any managed classes

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

		// Gmail tools
		public string getNewUserName(AppsService service, string studentID, string firstName, string midName, string lastName, int i, bool complete)
		{
			// this could really get screwed if there are enough duplicates it will be only do first.m.last f.middle.last
			// recursive function depends on the people to have longer names and different middle names hopefully
			string returnvalue = "";
			if (complete == false)
			{
				// ouDN = "OU=fakeou,DC=mydomain,DC=com
				int r = midName.Length;
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
						 returnvalue = firstName.Substring(0, (i - r)) + "." + midName + "." + lastName;
					}
				}

				try
				{
					service.CreateNickname(studentID, returnvalue);
				}
				catch
				{
					i++;
					getNewUserName(service, studentID, firstName, midName, lastName, i, complete);
				}
			}
			return returnvalue;
		}
		public void Create_Gmail_Users(AppsService service, GmailUsers gusersyn, SqlDataReader users, LogFile log)
		{
			// Takes the SQLDataReader and creates all users in the reader

			string studentID = "";
            string first_name = "";
			string last_name = "";
			string middle_name = "";
			string midI = "";
			string password = "";
			string alias = "";
			string userNickName = "";
			while (users.Read())
			{

				try
				{
					studentID = System.Web.HttpUtility.UrlEncode(users[gusersyn.User_StuID].ToString()).Replace("+", " ").Replace("*", "%2A");
					first_name = System.Web.HttpUtility.UrlEncode(users[gusersyn.User_Fname].ToString()).Replace("+", " ").Replace("*", "%2A");
					middle_name = System.Web.HttpUtility.UrlEncode(users[gusersyn.User_Mname].ToString()).Replace("+", " ").Replace("*", "%2A");
					last_name = System.Web.HttpUtility.UrlEncode(users[gusersyn.User_Lname].ToString()).Replace("+", " ").Replace("*", "%2A");
					password = System.Web.HttpUtility.UrlEncode(users[gusersyn.User_password].ToString()).Replace("+", " ").Replace("*", "%2A");
					//Create a new user.
					UserEntry insertedEntry = service.CreateUser(studentID, first_name, last_name, password);
					try
					{
						
					}
					catch
					{
						alias = " alias creation failed";
					}

					log.transactions.Add("Added Gmail user " + studentID + alias);
				}
				catch
				{
				}
			}

		}
		public DataTable Get_Gamil_Users(AppsService service, GmailUsers gusersyn, string table)
		{
			// nicknames will have to dealt with seperately
			DataTable returnvalue = new DataTable();
			DataRow row;

			returnvalue.TableName = table;

			int i = 0;
			int count = 0;
			returnvalue.Columns.Add();
			returnvalue.Columns.Add();
			returnvalue.Columns.Add();

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
				   row[0] = (userEntry.Name.GivenName.ToString());
				   row[1] = (userEntry.Name.FamilyName.ToString());
				   row[2] = (userEntry.Login.UserName.ToString());
			       
				   //userList.Add(userEntry.Login.UserName.ToString());

				   returnvalue.Rows.Add(row);
				   row = returnvalue.NewRow();
			   }
			}
			catch
			{
			  // result.AppendText("failed authentication");
			}
			return returnvalue;
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
		public void ExecuteGroupSync(GroupSynch groupsyn, ToolSet tools, LogFile log)
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
				// make the list of fields for the sql to check when updating note these fields must be in the same order as the AD update keys
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
			// we didn't find any records in AD so there is no need for the Update or delete logic to run
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
		public void ExecuteUserSync(UserSynch usersyn, ToolSet tools, LogFile log)
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

			ArrayList completeSqlKeys = new ArrayList();
			ArrayList completeADKeys = new ArrayList();
			ArrayList adUpdateKeys = new ArrayList();
			ArrayList sqlUpdateKeys = new ArrayList();
			ArrayList extraFieldsToReturn = new ArrayList();
			ArrayList fields = new ArrayList();
			Dictionary<string, string> userObject = new Dictionary<string, string>();
			SqlConnection sqlConn = new SqlConnection("Data Source=" + usersyn.DataServer + ";Initial Catalog=" + usersyn.DBCatalog + ";Integrated Security=SSPI;");


			string sqlUsersTable = "#sqlusersTable";
			string adUsersTable = "#ADusersTable";
			//SqlDataReader sqlusers;
			SqlCommand sqlComm;
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
					", RTRIM(" + usersyn.User_city + ") AS l" +
					", RTRIM(" + usersyn.User_Zip + ") AS postalcode" +
					", RTRIM(" + usersyn.User_password + ") AS password" +
					sqlForCustomFields +
					" INTO " + sqlUsersTable + " FROM " + usersyn.User_dbTable +
					" WHERE " + usersyn.User_where, sqlConn);
			}
			sqlComm.ExecuteNonQuery();

			// set up fields to pull back from SQL  
			// the custom ones are set up previously in the loop above while generating the sql statement
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


			// go grab all the users from AD
			adUsers = tools.EnumerateUsersInOUDataTable(usersyn.BaseUserOU, completeADKeys, adUsersTable);
			if (adUsers.Rows.Count > 0)
			{
				// make the temp table for ou comparisons
				tools.Create_Table(adUsers, adUsersTable, sqlConn);



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
				add = tools.QueryNotExists(sqlUsersTable, adUsersTable, sqlConn, "sAMAccountName", adUsers.Columns[0].ColumnName);


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

				//debugReader.Close();
				//MessageBox.Show("table " + adUsersTable + "\n " + debugFieldCount + " fields \n sample data" + debug);

				tools.CreateUserAccount(usersyn.UserHoldingTank, add, usersyn.UniversalGroup, usersyn, log);
				add.Close();

				delete = tools.QueryNotExists(sqlUsersTable, adUsersTable, sqlConn, usersyn.User_sAMAccount, completeADKeys[0].ToString());
				// delete groups in AD
				while (delete.Read())
				{
					tools.DeleteUserAccount((string)delete[0], (string)delete[1], log);
					// log.transactions.Add("User removed ;" + (string)delete[adUpdateKeys[1].ToString()].ToString().Trim()); 
				}
				delete.Close();

				// add the extra fields in form ".field ,"
				extraFieldsToReturn.Add(adUsersTable + ".distinguishedname ,");

				update = tools.CheckUpdate(sqlUsersTable, adUsersTable, usersyn.User_sAMAccount, "CN", sqlUpdateKeys, adUpdateKeys, extraFieldsToReturn, sqlConn);

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
			}
			sqlConn.Close();
		}
	}
	public class ObjectADGoogleSync
	{
		public void EmailNicknameSync()
		{

		}
		public void EmailUsersSync(GmailUsers gusersyn)
		{
			// appservice variables will come from a config designed ot hold its data (sql and Gmail login)
			AppsService service = new AppsService(gusersyn.Admin_domain, gusersyn.Admin_user, gusersyn.Admin_user);

		}
	}
}
