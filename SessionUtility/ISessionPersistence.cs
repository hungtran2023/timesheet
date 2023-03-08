using System;

namespace MSDN
{
	/// <summary>
	/// Summary description for ISessionPersistence.
	/// </summary>
	public interface ISessionPersistence
	{
		mySession LoadSession(string key, string dsn, int timeOut);
		void SaveSession(string key, string dsn, mySession Session, bool IsNewSession);
		string GenerateKey();
		string SessionID{ get;}
	}
}
