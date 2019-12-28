# PowerPointService [![NuGet version](https://badge.fury.io/nu/RandomSolutions.PowerPointService.svg)](http://badge.fury.io/nu/RandomSolutions.PowerPointService)
PowerPoint presentation from template and model

## Example

*Create from template*
```C#
var powerPointService = new RandomSolutions.PowerPointService();
var template = File.ReadAllBytes(@".\template.pptx");
var result = _powerPointService.CreateFromTemplate(template, (i, len) => new {
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

![](/Test/Images/example01.png)

*Slide/row iterations*
```C#
var result = _powerPointService.CreateFromTemplate(template, (i, len) => 
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

![](/Test/Images/example02.png)


[template.pptx](Test/Presentations/template_source.pptx)
[result.pptx](Test/Presentations/template_result.pptx)

[More examples in the test console application...](Test/CoreConsoleApp/Program.cs)

