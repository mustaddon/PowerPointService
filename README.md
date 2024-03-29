# PowerPointTool [![NuGet version](https://badge.fury.io/nu/PowerPointTool.svg)](http://badge.fury.io/nu/PowerPointTool)
PowerPoint presentation from template and model

### Example #1 - Create from template

```C#
var ppTool = new PPTool();
var template = File.ReadAllBytes(@".\template.pptx");
var result = ppTool.CreateFromTemplate(template, ctx => new {
  Title = "Example",
  Created = DateTimeOffset.Now,
  User = new { 
    Name = "TestName", 
    IsActive = true,
    Evaluation = 1000000,
  },
  Items = Enumerable.Range(1, 5).Select(x => $"item#{x}"),
});
```

![](https://raw.githubusercontent.com/mustaddon/PowerPointService/master/Test/Images/example01.png)

### Example #2 - Iterator for slides/rows

```C#
var result = ppTool.CreateFromTemplate(template, ctx => 
  Enumerable.Range(1, 3).Select(x => new {
    CompanyName = $"Company #{x}",
    Employees = Enumerable.Range(1, _rnd.Next(4, 12)).Select(xx => new {
        Name = $"Employee #{xx}",
        Email = $"emp{xx}@company{x}.test",
        Birthday = new DateTime(_rnd.Next(1980, 2000), _rnd.Next(1, 12), 1).AddDays(_rnd.Next(0, 30)),
    }),
  })
);
```

![](https://raw.githubusercontent.com/mustaddon/PowerPointService/master/Test/Images/example02.png)

Presentations: 
[template.pptx](https://raw.githubusercontent.com/mustaddon/PowerPointService/master/Test/Presentations/template_source.pptx?raw=true) ,
[result.pptx](https://raw.githubusercontent.com/mustaddon/PowerPointService/master/Test/Presentations/template_result.pptx?raw=true)

[Program.cs](https://github.com/mustaddon/PowerPointService/blob/master/Test/ConsoleApp/Program.cs)

