using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var t = new ATest.DBClass();
            var errMsg = string.Empty;
            t.set_TestContent("C#");
            var ret = t.Add(ref errMsg);
            Console.WriteLine(ret == 1 ? "新增成功" : errMsg);

            var rest = t.Query("select * from test", ref errMsg);

            while (!rest.EOF)
            {
                Console.WriteLine(rest.Fields["id"].Value + "  " 
                    + rest.Fields["content"].Value);
                rest.MoveNext();
            }
            t = null;
            Console.ReadKey();
        }
    }
}
