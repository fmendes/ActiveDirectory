using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;
using System.Text;
using System.DirectoryServices;

/// <summary>
/// Summary description for NetworkUsersMapping
/// </summary>
public class NetworkUsersMapping
{
	private static string m_sDomain;
	private static string sLDAPPath = System.Configuration.ConfigurationManager.AppSettings["LDAPPath"];
	private static string sUID = System.Configuration.ConfigurationManager.AppSettings["LDAPUID"];
	private static string sPwd = System.Configuration.ConfigurationManager.AppSettings["LDAPPwd"];
	private static string sObjectClass = System.Configuration.ConfigurationManager.AppSettings["LDAPUserObjectClass"];
	private static string sLDAPDomains = System.Configuration.ConfigurationManager.AppSettings["LDAPDomains"];
	public static string strDSN = ConfigurationManager.ConnectionStrings["Goaltender2008ConnectionString"].ToString();

	public static Label lblError, lblRowCount;
	protected static DataTable objTable, objTableUserMaint;
	protected static DataTable objDomainTable;

	public NetworkUsersMapping()
	{
		//
		// TODO: Add constructor logic here
		//
	}
	static NetworkUsersMapping()
	{
		CreateNetworkUserTable();

		objDomainTable = new DataTable( "NetworkDomains" );
		objDomainTable.Columns.Add( new DataColumn( "Domain", typeof( System.String ) ) );

	}

	private static void CreateNetworkUserTable()
	{
		objTable = new DataTable( "NetworkUsers" );

		objTable.Columns.Add( new DataColumn( "ApplicationUser_ID", typeof( System.String ) ) );
		objTable.Columns.Add( new DataColumn( "UserDomain", typeof( System.String ) ) );
		objTable.Columns.Add( new DataColumn( "Username", typeof( System.String ) ) );
		objTable.Columns.Add( new DataColumn( "LastName", typeof( System.String ) ) );
		objTable.Columns.Add( new DataColumn( "FirstName", typeof( System.String ) ) );
		objTable.Columns.Add( new DataColumn( "Phone", typeof( System.String ) ) );
		objTable.Columns.Add( new DataColumn( "Email", typeof( System.String ) ) );
		objTable.Columns.Add( new DataColumn( "FullName", typeof( System.String ) ) );
	}

	public static DataTable SelectDomains()
	{
		string ldapdomains = System.Configuration.ConfigurationManager.AppSettings[ "LDAPDomains" ].ToString();
		string[] Domains = ldapdomains.Split( new char[] { ';' } );

		objDomainTable.Rows.Clear();
		foreach ( string strDomain in Domains )
			objDomainTable.Rows.Add( new object[] { strDomain } );

		return objDomainTable;
	}

	public static DataTable Select(string strDomain)
	{
		string sLastNameSearch = "*";
		string sFirstNameSearch = "*";

		CreateNetworkUserTable();
		objTable.Rows.Clear();

		return LookForUserInDomain(strDomain, sLastNameSearch, sFirstNameSearch);
	}

	public static DataTable SelectUser(string strDomain, string sLastNameSearch, string sFirstNameSearch)
	{
		CreateNetworkUserTable();
		objTable.Rows.Clear();

		return LookForUserInDomain(strDomain, sLastNameSearch, sFirstNameSearch);
	}

	private static DataTable LookForUserInDomain(string strDomain, string sLastNameSearch, string sFirstNameSearch)
	{
		if (sUID == "") sUID = null;
		if (sPwd == "") sPwd = null;

		string serverName = System.Configuration.ConfigurationManager.AppSettings[strDomain].ToString();
		sLDAPPath = "LDAP://" + serverName + "/DC=" + strDomain + ",DC=root01,DC=org";

		objTable = GetLDAPUserInfo(sLDAPPath, sUID, sPwd, sObjectClass, sLastNameSearch, sFirstNameSearch);

		return objTable;
	}

	public static DataTable SelectUsersFromAllDomains()
	{
		string sLastNameSearch = "*";
		string sFirstNameSearch = "*";

		if (sUID == "") sUID = null;
		if (sPwd == "") sPwd = null;

		CreateNetworkUserTable();
		objTable.Rows.Clear();

		//Search in all the domains
		string ldapdomains = System.Configuration.ConfigurationManager.AppSettings["LDAPDomains"].ToString();
		string[] Domains = ldapdomains.Split(new char[] { ';' });

		for (int i = 0; i < Domains.Length; i++)
		{
			string domainName = Domains[i];

			objTable = LookForUserInDomain(domainName, sLastNameSearch, sFirstNameSearch);

			//string serverName = System.Configuration.ConfigurationManager.AppSettings[domainName].ToString();
			//sLDAPPath = "LDAP://" + serverName + "/DC=" + domainName + ",DC=root01,DC=org";
			////					gs.Domain = domainName;
			////					gs.LDAPPath = "LDAP://" + serverName + "/DC=" + domainName + ",DC=root01,DC=org";

			//objTable = GetLDAPUserInfo(sLDAPPath, sUID, sPwd, sObjectClass, sLastNameSearch, sFirstNameSearch);
		}

		return objTable;
	}


	public static DataTable LookForUserInAllDomains(string sLastNameSearch, string sFirstNameSearch)
	{
		if (sUID == "") sUID = null;
		if (sPwd == "") sPwd = null;

		CreateNetworkUserTable();
		objTable.Rows.Clear();

		////Search in all the domains
		//string ldapdomains = System.Configuration.ConfigurationManager.AppSettings["LDAPDomains"].ToString();
		//string[] Domains = ldapdomains.Split(new char[] { ';' });

		//for (int i = 0; i < Domains.Length; i++)
		//{
		//    string domainName = Domains[i];

		//    objTable = LookForUserInDomain(domainName, sLastNameSearch, sFirstNameSearch);

		//}

		string sFilter = String.Format("(|(&(objectClass=User)(givenname={0})(sn={1})))", sFirstNameSearch, sLastNameSearch);

		// collect inactive users in all the domains
		string[] sDomains = sLDAPDomains.Split(new char[] { ';' });
		for (int i = 0; i < sDomains.Length; i++ )
		{
			string sDomainName = sDomains[ i ];
			string sServerName = System.Configuration.ConfigurationManager.AppSettings[sDomainName].ToString();
			string sLDAPPath = "LDAP://" + sServerName + "/DC=" + sDomainName + ",DC=root01,DC=org";

			DirectoryEntry objRootDE = new DirectoryEntry(sLDAPPath, sUID, sPwd, AuthenticationTypes.Secure);
			DirectorySearcher objDS = new DirectorySearcher(objRootDE);

			objDS.Filter = sFilter;
			objDS.ReferralChasing = ReferralChasingOption.None;
			objDS.PropertiesToLoad.Add("userAccountControl");
			objDS.PropertiesToLoad.Add("SAMAccountName");
			objDS.PropertiesToLoad.Add("givenName");
			objDS.PropertiesToLoad.Add("sn");
			objDS.PropertiesToLoad.Add("TelephoneNumber");
			objDS.PropertiesToLoad.Add("mail");

			SearchResultCollection objSRC = null;
			try
			{
				objSRC = objDS.FindAll();
			}
			catch (Exception excpt)
			{
				if (excpt.Message.IndexOf("The server is not operational.") < 0)
					throw;
			}

			if (objSRC == null)
				continue;

			foreach (SearchResult objSR in objSRC)
			{
				int iInactiveFlag	= Convert.ToInt32(objSR.Properties["userAccountControl"][0]);
				string sUserId		= objSR.Properties["SAMAccountName"][0].ToString();
				string sFirstName	= objSR.Properties["givenName"][0].ToString();
				string sLastName	= objSR.Properties["sn"][0].ToString();

				string sPhone	= "";
				string sEmail	= "";

				if (objSR.Properties["TelephoneNumber"].Count > 0)
					sPhone	= objSR.Properties["TelephoneNumber"][0].ToString();

				if( objSR.Properties["mail"].Count > 0 )
					sEmail	= objSR.Properties["mail"][0].ToString();

				iInactiveFlag = iInactiveFlag & 0x0002;
				if (iInactiveFlag <= 0)
				{
					// add name, username, phone and email to the table, if active
					DataRow objRow = objTable.NewRow();

					objRow["LastName"] = sLastName;
					objRow["FirstName"] = sFirstName;
					objRow["Username"] = sUserId;
					objRow["UserDomain"] = sDomainName;
					objRow["Phone"] = sPhone;
					objRow["Email"] = sEmail;

					objTable.Rows.Add( objRow );

					continue;
				}
			}

			objSRC.Dispose();
			objDS.Dispose();
			objRootDE.Close();
			objRootDE.Dispose();
		}

		return objTable;

	}

	public static DataView SelectInactiveUsers()
	{
		StringBuilder strbList = new StringBuilder(""), strbSQLSelect = new StringBuilder("") ;

		CreateApplicationUserTable( "Inactive" );

		DataTable objUserDomain = new DataTable();
		CreateUserDomainTable(objUserDomain);

		// collect inactive users in all the domains
		string[] sDomains = sLDAPDomains.Split(new char[] { ';' });

		for (int i = 0; i < sDomains.Length; i++)
		{
			string sDomainName = sDomains[i];
			
			SearchResultCollection objSRC = RetrieveAllNetworkUsersFromLDAP(sDomainName);

			if (objSRC == null)
				continue;

			strbSQLSelect = new StringBuilder("");
			strbSQLSelect.AppendFormat(@"SELECT ApplicationUser_ID, UserDomain, Username, FirstName, LastName, 
							Phone, Email, Division_ID, DateCreated, DateUpdated, UserRole_ID, NewUser, 
							ReportDate, COALESCE( TerminatedFlag, CONVERT( bit, 0 ) ) AS TerminatedFlag, 
							'{0}' AS InactiveDomain 
						FROM dbo.ApplicationUser 
						WHERE UserName IN ( ", sDomainName);

			strbList = new StringBuilder("");

			// copy the results from the network to a DataTable and 
			// concatenate a list of inactive user ids found on the network
			foreach (SearchResult objSR in objSRC)
			{
				int iInactiveFlag = Convert.ToInt32(objSR.Properties["userAccountControl"][0]);
				string sUserId = objSR.Properties["SAMAccountName"][0].ToString();

				// skip active accounts
				iInactiveFlag = iInactiveFlag & 0x0002;
				if (iInactiveFlag <= 0)
				{
					// save active user x domain for later comparison
					DataRow objRow = objUserDomain.NewRow();
					objRow["Username"] = sUserId;
					objRow["Domain"] = sDomainName;
					objUserDomain.Rows.Add(objRow);

					continue;
				}
				
				// user is inactive in domain, add inactive user to the list
				strbList.AppendFormat("'{0}', ", sUserId);
				
				//objSR.Properties[ "SAMAccountName" ][ 0 ]
				//AddObjectToTable(sr, sLDAPUserObjectClass);
			}

			objSRC.Dispose();

			// remove the last comma
			if (strbList.Length > 2)
				strbList.Remove(strbList.Length - 2, 2);

			strbSQLSelect.Append(strbList.ToString());
			strbSQLSelect.Append(" ) ");

			// retrieve the inactive users for the domain
			SQLHelper.FillTable(strbSQLSelect.ToString(), ref objTableUserMaint, lblError);
		}

		// remove users that are inactive in one domain but active in another
		foreach (DataRow objRow in objTableUserMaint.Rows)
		{
			string sUser = objRow["Username"].ToString();
			string sDomain = objRow["UserDomain"].ToString();
			string sCriteria = String.Format("Username = '{0}' AND Domain = '{1}'", sUser, sDomain);

			// check whether the user found in the database is 
			// in the same domain as in the network
			// if it is, exclude it
			DataView objActiveUserDV = new DataView(objUserDomain,
						sCriteria, "", DataViewRowState.CurrentRows);
			if (objActiveUserDV.Count > 0)
				objRow["ExcludeFlag"] = true;
			else
				objRow["ExcludeFlag"] = false;
		}

		DataView objDV = new DataView(objTableUserMaint, "ExcludeFlag = false", "FirstName, LastName, UserName", DataViewRowState.CurrentRows);

		if (lblRowCount != null)
			lblRowCount.Text = String.Format("{0} Inactive users found / {1} Users Active in the Network",
				objDV.Count, objUserDomain.Rows.Count);

		return objDV;
	}

	private static void CreateApplicationUserTable( string strType )
	{
		objTableUserMaint = new DataTable("ApplicationUser");
		objTableUserMaint.Reset();
		objTableUserMaint.Columns.Add("ApplicationUser_ID", typeof(int));
		objTableUserMaint.Columns.Add("UserDomain", typeof(string));
		objTableUserMaint.Columns.Add("Username", typeof(string));
		objTableUserMaint.Columns.Add("FirstName", typeof(string));
		objTableUserMaint.Columns.Add("LastName", typeof(string));
		objTableUserMaint.Columns.Add("Phone", typeof(string));
		objTableUserMaint.Columns.Add("Email", typeof(string));
		objTableUserMaint.Columns.Add("Division_ID", typeof(int));
		objTableUserMaint.Columns.Add("DateCreated", typeof(DateTime));
		objTableUserMaint.Columns.Add("DateUpdated", typeof(DateTime));
		objTableUserMaint.Columns.Add("UserRole_ID", typeof(int));
		objTableUserMaint.Columns.Add("NewUser", typeof(int));
		objTableUserMaint.Columns.Add("ReportDate", typeof(DateTime));
		objTableUserMaint.Columns.Add("TerminatedFlag", typeof(bool));

		if( strType.Equals( "Inactive" ) )
			objTableUserMaint.Columns.Add("InactiveDomain", typeof(string));
		else
			objTableUserMaint.Columns.Add("NewDomain", typeof(string));

		objTableUserMaint.Columns.Add("ExcludeFlag", typeof(bool));
	}

	public static void DeleteUser(int ApplicationUser_ID)
	{
		string	strSQLSelect	= string.Format( "DELETE dbo.ApplicationUser WHERE ApplicationUser_ID = {0}", ApplicationUser_ID );
		SQLHelper.RunSQLCommand( strSQLSelect, lblError, false);
	}

	public static DataView SelectUsersWithWrongDomain()
	{
		StringBuilder strbList = new StringBuilder(""), strbSQLSelect = new StringBuilder("");

		CreateApplicationUserTable( "Active" );

		DataTable objUserDomain	= new DataTable();
		CreateUserDomainTable( objUserDomain );

		// collect users in all the domains
		string[] sDomains = sLDAPDomains.Split(new char[] { ';' });

		for (int i = 0; i < sDomains.Length; i++)
		{
			string sDomainName = sDomains[i];

			SearchResultCollection objSRC = RetrieveAllNetworkUsersFromLDAP( sDomainName );

			if (objSRC == null)
				continue;

			// create SQL to retrieve users in the database with a different domain than
			// the one from the network
			strbSQLSelect = new StringBuilder("");
			strbSQLSelect.AppendFormat(@"SELECT ApplicationUser_ID, UserDomain, Username, FirstName, LastName, 
							Phone, Email, Division_ID, DateCreated, DateUpdated, UserRole_ID, NewUser, 
							ReportDate, COALESCE( TerminatedFlag, CONVERT( bit, 0 ) ) AS TerminatedFlag, 
							'{0}' AS NewDomain 
						FROM dbo.ApplicationUser outau 
						WHERE NOT EXISTS ( SELECT au.UserDomain 
											FROM dbo.ApplicationUser AS au 
											WHERE au.UserDomain = '{0}' 
												AND au.UserName = outau.UserName ) 
							AND outau.UserName IN ( ", sDomainName);
			// removed: AND outau.UserDomain <> '{0}' 
			strbList = new StringBuilder("");

			// copy the results from the network to a DataTable and 
			// concatenate a list of user ids found on the network
			foreach (SearchResult objSR in objSRC)
			{
				int iInactiveFlag = Convert.ToInt32(objSR.Properties["userAccountControl"][0]);
				iInactiveFlag = iInactiveFlag & 0x0002;

				// skip inactive accounts
				if (iInactiveFlag > 0)
					continue;

				// get the user id
				string sUserId = objSR.Properties["SAMAccountName"][0].ToString();

				// save active user x domain for later comparison
				DataRow	objRow	= objUserDomain.NewRow();
				objRow["Username"]	= sUserId;
				objRow["Domain"]	= sDomainName;
				objUserDomain.Rows.Add(objRow);

				strbList.AppendFormat("'{0}', ", SQLHelper.ReplaceApostrophes(sUserId));
			}

			objSRC.Dispose();

			// remove the last comma
			if (strbList.Length > 2)
				strbList.Remove(strbList.Length - 2, 2);

			strbSQLSelect.Append(strbList.ToString());
			strbSQLSelect.Append(" ) ");

			// retrieve the users with incorrect domain 
			SQLHelper.FillTable(strbSQLSelect.ToString(), ref objTableUserMaint, lblError);
		}

		// remove users that are active in more than one domain
		foreach(DataRow objRow in objTableUserMaint.Rows)
		{
			string sUser = objRow["Username"].ToString();
			string sDomain = objRow["UserDomain"].ToString();
			string sCriteria = String.Format("Username = '{0}' AND Domain = '{1}'", sUser, sDomain );

			// check whether the user found in the database is 
			// in the same domain as in the network
			// if it is, exclude it
			DataView objActiveUserDV = new DataView(objUserDomain,
						sCriteria, "", DataViewRowState.CurrentRows);
			if( objActiveUserDV.Count > 0 )
				objRow[ "ExcludeFlag" ] = true;
			else
				objRow["ExcludeFlag"] = false;
		}

		DataView objDV = new DataView(objTableUserMaint, "ExcludeFlag = false", "FirstName, LastName, UserName", DataViewRowState.CurrentRows);

		if (lblRowCount != null)
			lblRowCount.Text = String.Format( "{0} Irregular records found / {1} Users Active in the Network", 
				objDV.Count, objUserDomain.Rows.Count );

		return objDV;
	}

	private static SearchResultCollection RetrieveAllNetworkUsersFromLDAP(string sDomainName)
	{
		string sServerName = System.Configuration.ConfigurationManager.AppSettings[sDomainName].ToString();
		string sLDAPPath = "LDAP://" + sServerName + "/DC=" + sDomainName + ",DC=root01,DC=org";

		DirectoryEntry objRootDE = new DirectoryEntry(sLDAPPath, sUID, sPwd, AuthenticationTypes.Secure);
		DirectorySearcher objDS = new DirectorySearcher(objRootDE);

		objDS.Filter = "(|(&(objectClass=User)(givenname=*)(sn=*)))";
		objDS.ReferralChasing = ReferralChasingOption.None;
		objDS.PropertiesToLoad.Add("userAccountControl");
		objDS.PropertiesToLoad.Add("SAMAccountName");

		SearchResultCollection objSRC = null;
		try
		{
			objSRC = objDS.FindAll();
		}
		catch (Exception excpt)
		{
			if (excpt.Message.IndexOf("The server is not operational.") < 0)
				throw;
		}

		objDS.Dispose();
		objRootDE.Close();
		objRootDE.Dispose();
		return objSRC;
	}

	private static void CreateUserDomainTable( DataTable objUserDomain )
	{
		objUserDomain.Reset();
		objUserDomain.Columns.Add("Username", typeof(string));
		objUserDomain.Columns.Add("Domain", typeof(string));
		objUserDomain.Columns.Add("FirstName", typeof(string));
		objUserDomain.Columns.Add("LastName", typeof(string));

	}

	public static void UpdateAllUsersWithWrongDomain()
	{
		DataView objDV = SelectUsersWithWrongDomain();

		// update all user domains to their network domains
		foreach (DataRowView objDRV in objDV)
		{
			int iAppUserId = Convert.ToInt32(objDRV["NewDomain"]);
			string strDomain = objDRV["ApplicationUser_ID"].ToString();

			UpdateUserDomain(iAppUserId, strDomain);
		}
	}

	public static void DeleteAllInactiveUsers()
	{
		DataView objDV = SelectInactiveUsers();

		// delete all inactive users
		foreach (DataRowView objDRV in objDV)
		{
			int iAppUserId = Convert.ToInt32(objDRV["ApplicationUser_ID"]);

			DeleteUser(iAppUserId );
		}

	}

	public static void UpdateUserDomain(int ApplicationUser_ID, string UserDomain)
	{
		string strSQLSelect = string.Format("UPDATE dbo.ApplicationUser SET UserDomain = '{0}' WHERE ApplicationUser_ID = {1}", UserDomain, ApplicationUser_ID);
		SQLHelper.RunSQLCommand(strSQLSelect, lblError, false);
	}



}
