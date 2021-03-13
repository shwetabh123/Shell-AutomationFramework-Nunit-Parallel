using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test.Helpers
{
   public class DatFileHelper
    {


        public string datFilePath;
        public string datFileName;

        public DatFileHelper(string datFilePath )
        {

            this.datFilePath = datFilePath;

        }


        public static void createdatfile(string datFilePath,string datFileName)
        {

            string file = datFilePath + "\\" + datFileName + ".dat";

            if(!(File.Exists(file) ))
            {
                using (StreamWriter sw=File.CreateText(file))
                {

                    sw.WriteLine("D|AN07|1|086||999999|KBSite 999|1|1||7077861008038000016|0824|20200101100000|PHP|1|100||26||1|1|090909|||||y|additional 1|additional 2|additional 3|additional 4");


                    sw.WriteLine("T|DX026_GFN_TRX_00001_000093_20200813_160000.dat|1|100");

                }


            }
        }

    }
}
