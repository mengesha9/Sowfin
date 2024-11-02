using System;
using System.Collections.Generic;
using System.Text;

namespace Sowfin.Model.Entities
{
 

    public class Interest_IValuation
    {
        public long Id { get; set; }
        public string? Interest_Income { get; set; }  
        public string? Interest_Expense { get; set; } 
        public long? InitialSetupId { get; set; }     
        public long UserId { get; set; }
        public DateTime? CreatedDate { get; set; }    
        public DateTime? ModifiedDate { get; set; }  
        public string? TableData { get; set; }        
        public int? Year { get; set; }               
    }


}
