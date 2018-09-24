using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace RhPro.Oodd.Utilities
{
    /// <summary>
    /// Maneja y extiende string
    /// </summary>
    public static class StringManager
    {
        public static string ToTitleCase(this string value)
        {
            return ToTitleCase(value, new List<char> { ' ' });
        }
      
        public static string ToTitleCase(this string value, List<char> separators)
         {
            string result = "";
            bool nextUpper = true; //first letter always upper case
     
            value = value.ToLower();//initialize all to lower case
     
            for (int charIndex = 0; charIndex < value.Length; charIndex++)
            {
               string nextChar = value[charIndex].ToString();
                if (nextUpper)
                {
                   nextChar = nextChar.ToUpper();
                }
     
                result += nextChar;
     
                if (separators.Any(c => c.Equals(value[charIndex])))//put next char to upper case
                    nextUpper = true;
                else
                    nextUpper = false;
     
            }

            return result.TrimEnd(' ').TrimStart(' ');
        }

    }
}
