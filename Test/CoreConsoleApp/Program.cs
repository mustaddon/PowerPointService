using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using RandomSolutions;

namespace CoreConsoleApp
{
    class Program
    {
        static PowerPointService _powerPointService = IoC.Container.GetInstance<PowerPointService>();
        static Random _rnd = new Random();

        static void Main(string[] args)
        {
            _testMergeSlides();
            _testInsertSlides();
            _testDeleteSlides();
            _testCreateFromTemplate();
        }

        static string _presentationsDir(string file) => Path.Combine(@"..\..\..\..\presentations\", file);

        static void _testMergeSlides()
        {
            var target = File.ReadAllBytes(_presentationsDir("merge_target.pptx"));
            var source = File.ReadAllBytes(_presentationsDir("merge_source.pptx"));
            var result = _powerPointService.MergeSlides(source, target, (i, slen, tlen) => i==0 ? 1 : -1);
            File.WriteAllBytes(_presentationsDir("merge_result.pptx"), result);
        }

        static void _testInsertSlides()
        {
            var target = File.ReadAllBytes(_presentationsDir("insert_target.pptx"));
            var source = File.ReadAllBytes(_presentationsDir("insert_source.pptx"));
            var result = _powerPointService.InsertSlides(source, target, 1, (i, len) => i > 0);
            File.WriteAllBytes(_presentationsDir("insert_result.pptx"), result);
        }

        static void _testDeleteSlides()
        {
            var source = File.ReadAllBytes(_presentationsDir("delete_source.pptx"));
            var result = _powerPointService.DeleteSlides(source, (i, len) => i > 0);
            File.WriteAllBytes(_presentationsDir("delete_result.pptx"), result);
        }

        static void _testCreateFromTemplate()
        {
            var template = File.ReadAllBytes(_presentationsDir("template_source.pptx"));

            var result = _powerPointService.CreateFromTemplate(template, (i, len) =>
            {
                if (i == len - 1)
                    return Enumerable.Range(1, 3).Select(x => new
                    {
                        CompanyName = $"Company #{x}",
                        Employees = Enumerable.Range(1, _rnd.Next(4, 12)).Select(xx => new
                        {
                            Name = $"Employee #{xx}",
                            Email = $"emp{xx}@company{x}.test",
                            Birthday = new DateTime(_rnd.Next(1980, 2000), _rnd.Next(1, 12), 1).AddDays(_rnd.Next(0, 30)),
                        }),
                    });

                return new
                {
                    Title = "Example",
                    Created = DateTimeOffset.Now,
                    User = new { 
                        Name = "TestName", 
                        IsActive = true,
                        Evaluation = 1000000,
                    },
                    Items = Enumerable.Range(1, 5).Select(x => $"item#{x}"),
                };
            });

            File.WriteAllBytes(_presentationsDir("template_result.pptx"), result);
        }

    }
}
