using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class OneVisionAgency_GE_Properties
    {
        public int Sl_Number
        {
              get; set;
        }
        
        public string Message_Name
        {
              get; set;
        }

        public string Application_Name
        {
            get; set;
        }
        public string Success_Archive_FolderLocation
        {
            get; set;
        }
        public string Direction
        {
            get; set;
        }
        public string Priority
        {
            get; set;
        }
        public int Files_Count
        {
            get; set;
        }
        public int Failed_Count
        {
            get; set;

        }
        public string Failed_Reason
        {
            get; set;

        }
        public DateTime Date_to_fix
        {
            get;set;
        }
    }
}
