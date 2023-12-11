# ThinkSharp.WordTemplating

[![NuGet](https://img.shields.io/nuget/v/ThinkSharp.WordTemplating.svg)](https://www.nuget.org/packages/ThinkSharp.WordTemplating/)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=MSBFDUU5UUQZL)

## Introduction

ThinkSharp.WordTemplating is a lightweight templating engine for OpenXml (docx) documents with simple replacement capabilities based on DocumentFormat.OpenXml.

Usage:
## Installation

The library is basically, one class that can be copied and modified if needed. However, there is also a nuget version: [Nuget](https://www.nuget.org/packages/ThinkSharp.WordTemplating)

      Install-Package ThinkSharp.WordTemplating
      
## Usage

  var replacements = new Dictionary<string, object>(StringComparer.InvariantCultureIgnoreCase);
  
  replacements.Add("Replacement1", "Text For Replacement1");
  replacements.Add("Replacement2", "Text For Replacement2");
  replacements.Add("MyInt", 5);
  replacements.Add("MyDate", DateTime.Now);
  replacements.Add("MyDouble", Math.PI);
  replacements.Add("MyBoolean", true);
  replacements.Add("MyDateTimeOffset", DateTimeOffset.Now);
  replacements.Add("CustomType", new CustomType { FirstName = "Hugo", LastName = "Boss" } );
  replacements.Add("CustomTypeNoConverter", new CustomTypeNoConverter { FirstName = "Hugo", LastName = "Boss" });
  
  var engine = new TemplateEngine();
  
  // add custom converters (overwrites default converts)
  engine.AddReplacementConverter<bool>((val, format) => format == "YesNo" ? (val ? "Yes" : "No") : val.ToString());
  engine.AddReplacementConverter<CustomType>((val, format) => format == "LastFirst" ? $"{val.LastName}, {val.FirstName}" : val.ToString());
  
  File.Copy(@"Data\Test.docx", @"Data\Test_Replaced.docx", true);
  
  engine.ReplaceTemplate(@"Data\Test_Replaced.docx", replacements);
  
  public class CustomType
  {
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public override string ToString() => $"{FirstName} {LastName}";
  }
  
  public class CustomTypeNoConverter
  {
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public override string ToString() => $"{FirstName} {LastName}";
  }


## License

ThinkSharp.WordTemplating is released under [The MIT license (MIT)](LICENSE.TXT)

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/JanDotNet/ThinkSharp.WordTemplating/tags). 
    
## Donation
If you like ThinkSharp.WordTemplating and use it in your project(s), feel free to give me a cup of coffee :) 

[![paypal](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=MSBFDUU5UUQZL)
