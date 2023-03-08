using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization;
using System.Runtime.InteropServices;

namespace MSDN
{
	public class SessionPersistence : ISessionPersistence
	{
		private const string _SessionID = "mySession";
		private const string command = "select Data, Last_Accessed from SessionState where ID = @ID";
		private const string UpdateStatement = "update SessionState set Data = @Data, Last_Accessed = @LastAccessed where ID = @ID";
		private const string InsertStatement = "insert into SessionState values(@ID, @Data, @LastAccessed)";
		
		public SessionPersistence()
		{
		}

		private Byte[] Serialize(mySession Session)
		{
			if (Session == null) return null;

			Stream stream = null;
			Byte[] state = null;

			try
			{
				IFormatter formatter = new BinaryFormatter();
				stream = new MemoryStream();
				formatter.Serialize(stream, Session);
				state = new Byte[stream.Length];
				stream.Position = 0;
				stream.Read(state, 0, (int)stream.Length);
				stream.Close();
			}
			finally
			{
				if (stream != null)
					stream.Close();
			}

			return state;
		}

		private mySession Deserialize(Byte[] state)
		{
			if (state == null) return null;
			
			mySession Session = null;
			Stream stream = null;

			try
			{
				stream = new MemoryStream();
				stream.Write(state, 0, state.Length);
				stream.Position = 0;
				IFormatter formatter = new BinaryFormatter();
				Session = (mySession)formatter.Deserialize(stream);
			}
			finally
			{
				if (stream != null)
					stream.Close();
			}

			return Session;
		}

		public  mySession LoadSession(string key, string dsn, int SessionExpiration)
		{
			SqlConnection conn = new SqlConnection(dsn);
			SqlCommand LoadCmd = new SqlCommand();
			LoadCmd.CommandText = command;
			LoadCmd.Connection = conn;
			SqlDataReader reader = null;
			mySession Session = null;

			try
			{
				LoadCmd.Parameters.Add("@ID", new Guid(key));
				conn.Open();
				reader = LoadCmd.ExecuteReader();
				if (reader.Read())
				{
					DateTime LastAccessed = reader.GetDateTime(1).AddMinutes(SessionExpiration);
					if (LastAccessed >= DateTime.Now)
						Session = Deserialize((Byte[])reader["Data"]);
					else 
						Session = new mySession();
				}


			}
			finally
			{
				if (reader != null)
					reader.Close();
				if (conn != null)
					conn.Close();
			}
			
			return Session;
		}
		
		public void SaveSession(string key, string dsn, mySession Session, bool IsNewSession)
		{
			SqlConnection conn = new SqlConnection(dsn);
			SqlCommand SaveCmd = new SqlCommand();			
			SaveCmd.Connection = conn;
			
			try
			{


				if (IsNewSession)
					SaveCmd.CommandText = InsertStatement;
				else
					SaveCmd.CommandText = UpdateStatement;

				SaveCmd.Parameters.Add("@ID", new Guid(key));
				SaveCmd.Parameters.Add("@Data", Serialize(Session));
				SaveCmd.Parameters.Add("@LastAccessed", DateTime.Now.ToString());
		
				conn.Open();
				SaveCmd.ExecuteNonQuery();
			}
			finally
			{
				if (conn != null)
					conn.Close();
			}
		}

		public string GenerateKey()
		{
			return Guid.NewGuid().ToString();
		}

		public string SessionID
		{
			get
			{
				return _SessionID;
			}
		}
	}
}
