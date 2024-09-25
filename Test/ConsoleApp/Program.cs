using PowerPointTool;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            TestCreateFromTemplate();
            TestCreateFromTemplateStreams();
            TestUpdateSlides();
            TestMergeSlides();
            TestInsertSlides();
            TestDeleteSlides();
            TestSlideIndex();
        }

        static readonly PPTool _ppt = new();
        static readonly Random _rnd = new();

        static string PresDir(string file) => Path.Combine(@"..\..\..\..\presentations\", file);

        static void TestCreateFromTemplate()
        {
            var template = File.ReadAllBytes(PresDir("template_source.pptx"));

            var result = _ppt.CreateFromTemplate(template, ctx =>
            {
                if (ctx.SlideIndex == ctx.SlidesCount - 1) // last slide
                    return Enumerable.Range(1, 3).Select(x => new
                    {
                        CompanyName = $"Company #{x}",
                        Employees = Enumerable.Range(1, _rnd.Next(4, 12)).Select(xx => new
                        {
                            Name = $"Employee #{xx}",
                            Email = $"emp{xx}@company{x}.test",
                            Birthday = new DateOnly(_rnd.Next(1980, 2000), _rnd.Next(1, 12), 1).AddDays(_rnd.Next(0, 30)),
                        }),
                    });

                return new
                {
                    Title = "Example",
                    Created = DateTimeOffset.Now,
                    User = new
                    {
                        Name = "TestName",
                        IsActive = true,
                        Evaluation = 1000000,
                    },
                    Items = Enumerable.Range(1, 5).Select(x => $"item#{x}"),
                    LinkName = "TestLink",
                    PackageName = typeof(PPTool).FullName,
                    Html = "<p>text <b>text</b> text</p>"
                };
            });

            File.WriteAllBytes(PresDir("template_result.pptx"), result);
        }

        static void TestCreateFromTemplateStreams()
        {
            using var template = File.OpenRead(PresDir("template_source.pptx"));
            using var result = File.Create(PresDir("template_result_streams.pptx"));

            _ppt.CreateFromTemplate(result, template, ctx =>
            {
                if (ctx.SlideIndex == ctx.SlidesCount - 1) // last slide
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
                    User = new
                    {
                        Name = "TestName",
                        IsActive = true,
                        Evaluation = 1000000,
                    },
                    Items = Enumerable.Range(1, 5).Select(x => $"item#{x}"),
                    LinkName = "TestLink",
                    PackageName = typeof(PPTool).FullName,
                };
            });
        }

        static void TestUpdateSlides()
        {
            var template = File.ReadAllBytes(PresDir("template_source.pptx"));
            var jpg = File.ReadAllBytes(PresDir("img01.jpg"));

            var result = _ppt.UpdateSlides(template, ctx =>
            {
                if (ctx.SlideIndex == 0)
                    ctx.AddImage(jpg, new(0, 0, -ctx.SlideWidth / 4, ctx.SlideHeight / 2));

                if (ctx.SlideIndex == ctx.SlidesCount - 1) // last slide
                    ctx.ApplyModels(Enumerable.Range(1, 3).Select(x => new
                    {
                        Num = x,
                        CompanyName = $"Company #{x}",
                        Employees = Enumerable.Range(1, _rnd.Next(4, 12)).Select(xx => new
                        {
                            Name = $"Employee #{xx}",
                            Email = $"emp{xx}@company{x}.test",
                            Birthday = new DateTime(_rnd.Next(1980, 2000), _rnd.Next(1, 12), 1).AddDays(_rnd.Next(0, 30)),
                        }),
                    }), s => s.AddImage(jpg, new(ctx.SlideWidth * 7 / 8, 0, ctx.SlideWidth / 8, ctx.SlideHeight / 4)));
                else
                    ctx.ApplyModel(new
                    {
                        Title = "Example",
                        Created = DateTimeOffset.Now,
                        User = new
                        {
                            Name = "TestName",
                            IsActive = true,
                            Evaluation = 1000000,
                        },
                        Items = Enumerable.Range(1, 5).Select(x => $"item#{x}"),
                        LinkName = "TestLink",
                        PackageName = typeof(PPTool).FullName,
                        Html = "<p>text <b>text</b> text</p>"
                    });
            });

            File.WriteAllBytes(PresDir("update_result.pptx"), result);
        }

        static void TestSlideIndex()
        {
            using var template = File.OpenRead(PresDir("template_source.pptx"));
            var result = _ppt.SlideIndex(template, new Regex(@"CompanyName"));
            Console.WriteLine(result);
        }

        static void TestMergeSlides()
        {
            var target = File.ReadAllBytes(PresDir("merge_target.pptx"));
            var source = File.ReadAllBytes(PresDir("merge_source.pptx"));
            var result = _ppt.MergeSlides(source, target, ctx => ctx.SlideIndex == 0 ? 1 : -1);
            File.WriteAllBytes(PresDir("merge_result.pptx"), result);
        }

        static void TestInsertSlides()
        {
            var target = File.ReadAllBytes(PresDir("insert_target.pptx"));
            var source = File.ReadAllBytes(PresDir("insert_source.pptx"));
            var result = _ppt.InsertSlides(source, target, 1, ctx => ctx.SlideIndex > 0);
            File.WriteAllBytes(PresDir("insert_result.pptx"), result);
        }

        static void TestDeleteSlides()
        {
            var source = File.ReadAllBytes(PresDir("delete_source.pptx"));
            var result = _ppt.DeleteSlides(source, ctx => ctx.SlideIndex > 0);
            File.WriteAllBytes(PresDir("delete_result.pptx"), result);
        }
    }
}
