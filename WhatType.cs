using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataExcel
{
    class WhatType
    {
        public static bool IsInt(string value)
        {
            bool result = int.TryParse(value, out _);
           
            return result;
        }

        public static bool IsFloat(string value) 
        {
            bool result = float.TryParse(value, out _);
            return result;
        }

        public static int AddOne(int value)
        {
            return value + 1;
        }

    }
}
