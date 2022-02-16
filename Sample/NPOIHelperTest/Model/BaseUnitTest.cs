using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;

namespace NPOIHelperTest
{
    public class BaseUnitTest
    {
        public static void WriteLine<T>(T t)
        {
            var msg = JsonConvert.SerializeObject(t);
            WriteLine(msg);
        }

        public static void WriteLine<T>(IEnumerable<T> list)
        {
            foreach (var item in list)
            {
                var msg = JsonConvert.SerializeObject(item);
                WriteLine(msg);
            }
        }

        public static void WriteLine(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                var msg = JsonConvert.SerializeObject(row.ItemArray);
                WriteLine(msg);
            }
        }

        public static void WriteLine(string msg)
        {
            //Debug.WriteLine(msg);
            //Trace.WriteLine(msg);
            Console.WriteLine(msg);
        }

    }
}