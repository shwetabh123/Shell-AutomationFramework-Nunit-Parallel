using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace Test.Config
{
   public class Settings
    {
        public static int Timeout { get; set; }
        public BrowserType BrowserType { get; set; }
        public BrowserType LogPath { get; set; }
    }
}
