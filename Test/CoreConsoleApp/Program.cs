using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using RandomSolutions;

namespace CoreConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var modelFac = new Func<int, int, object>((x, len) =>
            {
                var model = new
                {
                    Request = new
                    {
                        Title = "TEST",
                        Items = Enumerable.Range(0, 2).Select(i => new { Name = $"name:{i}", Date = DateTimeOffset.Now.AddDays(i) }).ToArray()
                    },
                    test = $"TEXT:{x}",
                    html = "<html><!-- comment --><div>text <i>iii</i></div><ul><li>111</li><li>222</li></ul><a href=\"http://aaa.aaa\">aaa <b>bbb</b></a><html>",
                    b = Enumerable.Range(1, 3).Select(xx=> new { Title = $"user#{xx}" }),
                    c = new[] { "aaa", "bbb" },
                    d = DateTimeOffset.Now,
                    e = new[] { DateTimeOffset.Now, DateTimeOffset.Now.AddDays(1) },
                    num = "0.12345",
                    nums = new object[] { null, 12345.67890, 12345.67890f, 12345.67890m, 1234567890, 1234567890L },
                    rows = Enumerable.Range(0, 5).Select(i => new { col1 = $"col1:{i}", col2 = $"col2:{i}", col3 = $"col3:{i}" }),
                    bool_ = true,
                    bools = new bool?[] { true, false, null },
                };
                //return model;
                return x == 0 ? new[] { model, model } : new object[] { model, model, model } as object;
            });

            var powerPointService = IoC.Container.GetInstance<PowerPointService>();
            var data = powerPointService.CreateFromTemplate(File.ReadAllBytes(@"..\tmp.pptx"), modelFac);
            File.WriteAllBytes(@"..\test.pptx", data);

            Console.WriteLine("Hello World!");
        }
    }
}
