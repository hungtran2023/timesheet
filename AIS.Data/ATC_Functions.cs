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
    
    public partial class ATC_Functions
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public ATC_Functions()
        {
            this.ATC_Functions1 = new HashSet<ATC_Functions>();
            this.ATC_Permissions = new HashSet<ATC_Permissions>();
        }
    
        public int FunctionID { get; set; }
        public string Description { get; set; }
        public string Form { get; set; }
        public Nullable<int> GroupID { get; set; }
        public Nullable<byte> LevelOrder { get; set; }
        public bool fgUpdateable { get; set; }
        public bool fgAttribute { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ATC_Functions> ATC_Functions1 { get; set; }
        public virtual ATC_Functions ATC_Functions2 { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ATC_Permissions> ATC_Permissions { get; set; }
    }
}