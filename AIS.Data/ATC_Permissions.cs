//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace AIS.Data
{
    using System;
    using System.Collections.Generic;
    
    public partial class ATC_Permissions
    {
        public int GroupID { get; set; }
        public int FunctionID { get; set; }
        public Nullable<bool> Updateable { get; set; }
    
        public virtual ATC_Functions ATC_Functions { get; set; }
        public virtual ATC_Group ATC_Group { get; set; }
    }
}