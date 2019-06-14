using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace CreateExcelFile
{
    public class MessageElement : ConfigurationElement
    {
        [ConfigurationProperty("SerialNumber")]
        public int SerialNumber
        {
            get { return (int)this["SerialNumber"]; }
            set { this["SerialNumber"] = value; }
        }

        [ConfigurationProperty("MessageName")]
        public string MessageName
        {
            get { return (string)this["MessageName"]; }
            set { this["MessageName"] = value; }
        }

        [ConfigurationProperty("ApplicationName")]
        public string ApplicationName
        {
            get { return (string)this["ApplicationName"]; }
            set { this["ApplicationName"] = value; }
        }

        [ConfigurationProperty("ServerLocation")]
        public string ServerLocation
        {
            get { return (string)this["ServerLocation"]; }
            set { this["ServerLocation"] = value; }
        }

        [ConfigurationProperty("DirectionName")]
        public string DirectionName
        {
            get { return (string)this["DirectionName"]; }
            set { this["DirectionName"] = value; }
        }

        [ConfigurationProperty("FailedFileCountPath")]
        public string FailedFileCountPath
        {
            get { return (string)this["FailedFileCountPath"]; }
            set { this["FailedFileCountPath"] = value; }
        }
    }
}
