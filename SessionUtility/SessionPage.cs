using System;
using System.Xml;
using System.Collections.Specialized;
using System.Web;
using System.Configuration;

namespace MSDN
{
	/// <summary>
	/// Summary description for myPage.
	/// </summary>
	public class SessionPage : System.Web.UI.Page
	{
		private const int DEFAULT_TIMEOUT = 20;
		private const string CONNECTION_STRING  = "SessionDSN";
		private const string TIME_OUT = "SessionTimeOut";
		private HttpCookie cookie;
		private bool IsNewSession = false;
		private ISessionPersistence sessionPersistence = new SessionPersistence();

		public new mySession Session = null;
			
		private int SessionExpiration
		{
			get
			{
				if (ConfigurationSettings.AppSettings[TIME_OUT] != null)
					return Convert.ToInt32(ConfigurationSettings.AppSettings[TIME_OUT]);
				else
					return DEFAULT_TIMEOUT;
			}
		}

		private string dsn
		{
			get
			{
				return ConfigurationSettings.AppSettings[CONNECTION_STRING];
			}
		}

		private void InitializeComponent()
		{    
			cookie = this.Request.Cookies[sessionPersistence.SessionID];

			if (cookie == null)
			{
				Session = new mySession();
				CreateNewSessionCookie();
				IsNewSession = true;
			}
			else
				Session = sessionPersistence.LoadSession(Server.UrlDecode(cookie.Value).ToLower().Trim(), dsn, SessionExpiration);
				
			this.Unload += new EventHandler(this.PersistSession);
		}
		
		private void CreateNewSessionCookie()
		{
			cookie = new HttpCookie(sessionPersistence.SessionID, sessionPersistence.GenerateKey());
			this.Response.Cookies.Add(cookie);
		}
	
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: This call is required by the ASP.NET Web Form Designer.
			//
			InitializeComponent();
			base.OnInit(e);
		}

		private void PersistSession(Object obj, System.EventArgs arg)
		{
			sessionPersistence.SaveSession(Server.UrlDecode(cookie.Value).ToLower().Trim(), dsn, Session, IsNewSession);

		}
	}

}
