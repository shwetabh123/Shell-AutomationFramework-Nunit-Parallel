using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace Test.Helpers
{
   public class RandomNumberandStringGenerator
    {
        public static void randomnumberwithoutargument()
        {
        }
        public static string randomnumberwithoneargument(int length)
        {
            string b = "0123456789";
            var ran = new Random();
            string random = "";
            for (int i=0;i<length;i++)
            {
                int a = ran.Next(b.Length);
                random = random + b.ElementAt(a);
            }
            return random;
        }
        public static void randomnumberwithtwoargument()
        {
        }
        public static void randomstring()
        {
        }
        public static void randomalphanumericstringwithspecialcharacters()
        {
        }
    }
}
