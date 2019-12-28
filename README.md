# PowerPointService [![NuGet version](https://badge.fury.io/nu/RandomSolutions.PowerPointService.svg)](http://badge.fury.io/nu/RandomSolutions.PowerPointService)
PowerPoint presentation from template and model

## Example

*Simple usage*
```C#
var powerPointService = new RandomSolutions.PowerPointService();
var template = File.ReadAllBytes(@".\template.pptx");
var result = _powerPointService.CreateFromTemplate(template, (i, len) => new {
  Title = "Example",
  Created = DateTimeOffset.Now,
});
```

![](/Test/Images/example01.png)


[More examples in the test console application...](Test/CoreConsoleApp/Program.cs)

[Presentation sources...](Test/Presentations/)
