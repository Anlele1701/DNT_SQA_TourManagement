using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DAPM_TOURDL
{
    public class Compare
    {
        public string[] ConvertToArray(string[] parts)
        {
            string[] newString= new string[parts.Length];
            for (int j = 0; j < parts.Length; j++)
            {
                if (parts[j] == "null")
                {
                    newString[j] = "";
                }
                else
                {
                    newString[j] = parts[j];
                }
                Console.WriteLine(newString[j]);
            }
            return newString;
        }

        public bool CompareExpectedAndActual(string expected, string actual)
        {
            if (expected == actual) return true;
            else return false;
        }
    }
}