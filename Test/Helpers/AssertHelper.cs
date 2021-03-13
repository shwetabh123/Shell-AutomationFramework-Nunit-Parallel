using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test.Helpers
{
    public class AssertHelper
    {

        public static bool AreEqual(string expected, string actual)
        {
            try
            {
                Assert.AreEqual(expected, actual);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Assert.Fail("Strings are not matching");

                return false;

            }
        }


    }
}
