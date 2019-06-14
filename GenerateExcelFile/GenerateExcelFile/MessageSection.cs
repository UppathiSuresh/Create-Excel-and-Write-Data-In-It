using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace CreateExcelFile
{
   public class MessageSection: ConfigurationSection
    {

        [ConfigurationProperty("Messages", IsDefaultCollection = true)]
        public MessageCollections Messages
        {
            get { return (MessageCollections)this["Messages"]; }
            set { this["Messages"] = value; }
        }
    }
}
