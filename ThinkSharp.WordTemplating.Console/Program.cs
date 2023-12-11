using System.Runtime.CompilerServices;
using ThinkSharp.WordTemplating;

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