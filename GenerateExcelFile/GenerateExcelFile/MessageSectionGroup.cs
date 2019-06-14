using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace CreateExcelFile
{
   public class MessageSectionGroup: ConfigurationSectionGroup
    {

        [ConfigurationProperty("MessageSection", IsRequired = false)]
        public MessageSection ContextSettings
        {
            get { return (MessageSection)base.Sections["MessageSection"]; }
        }
    }
}
