using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace CreateExcelFile
{
    [ConfigurationCollection(typeof(MessageElement), AddItemName = "Message")]
   public class MessageCollections: ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new MessageElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((MessageElement)element).SerialNumber;
        }
    }
}
