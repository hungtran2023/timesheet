using AIS.Domain.Email;
using System.Collections;

namespace AIS.Domain.Common.Helper
{
    public class EmailTemplateHelper
    {
        public static readonly Hashtable PlaceHolderButton = new Hashtable
        {
            {"Manager", "#Manager#"},
            {"Requester","#Requester#"},
            {"DateFrom", "#DateFrom#"},
            {"DateTo", "#DateTo#"},
            {"Note",  "#Note#" },
             {"ProjectID",  "#ProjectID#" },
              {"EmailManager",  "#EmailManager#" },
              {"TimeClose",  "#TimeClose#" },
                 {"Server",  "#Server#" },
            {"LinkApproveRequestForManager",  "#LinkApproveRequestForManager#" },
        };

        public static string ReplaceHolder(string content, EmailReplaceHolderModel emailHolder)
        {
            string result = content;
            if (content == null) return "";
            if(emailHolder.Manager != null)
            {
                result = result.Replace(PlaceHolderButton["Manager"].ToString(), emailHolder.Manager);
            }
            if (emailHolder.Requester != null)
            {
                result = result.Replace(PlaceHolderButton["Requester"].ToString(), emailHolder.Requester);
            }
            if (emailHolder.Requester != null)
            {
                result = result.Replace(PlaceHolderButton["DateFrom"].ToString(), emailHolder.DateFrom);
            }
            if (emailHolder.Requester != null)
            {
                result = result.Replace(PlaceHolderButton["DateTo"].ToString(), emailHolder.DateTo);
            }
            if (emailHolder.Requester != null)
            {
                result = result.Replace(PlaceHolderButton["Note"].ToString(), emailHolder.Note);
            }
            if (emailHolder.ProjectID != null)
            {
                result = result.Replace(PlaceHolderButton["ProjectID"].ToString(), emailHolder.ProjectID);
            }

            if (emailHolder.EmailManager != null)
            {
                result = result.Replace(PlaceHolderButton["EmailManager"].ToString(), emailHolder.EmailManager);
            }
            if (emailHolder.Requester != null)
            {
                result = result.Replace(PlaceHolderButton["LinkApproveRequestForManager"].ToString(), emailHolder.LinkAprroveRequestForManager);
            }
            if (emailHolder.TimeClose != null)
            {
                result = result.Replace(PlaceHolderButton["TimeClose"].ToString(), emailHolder.TimeClose);
            }
            if (emailHolder.Server != null)
            {
                result = result.Replace(PlaceHolderButton["Server"].ToString(), emailHolder.Server);
            }
            return result;
        }
    }
}
