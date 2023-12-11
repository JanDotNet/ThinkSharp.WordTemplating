using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace ThinkSharp.WordTemplating
{
  public class TemplateEngine
  {
    private readonly Dictionary<Type, Func<object, CultureInfo, string, string>> converters = new Dictionary<Type, Func<object, CultureInfo, string, string>>();
    private readonly Regex regex = new Regex(@"\<(?<name>[\w_-]+)(:(?<format>[\s:\w._-]+))?>");
    private CultureInfo culture = new CultureInfo("de-DE");
    public TemplateEngine()
    {
      AddReplacementConverter<double>((val, format) => string.IsNullOrWhiteSpace(format) ? val.ToString() : val.ToString(format, culture));
      AddReplacementConverter<decimal>((val, format) => string.IsNullOrWhiteSpace(format) ? val.ToString() : val.ToString(format, culture));
      AddReplacementConverter<float>((val, format) => string.IsNullOrWhiteSpace(format) ? val.ToString() : val.ToString(format, culture));
      AddReplacementConverter<float>((val, format) => string.IsNullOrWhiteSpace(format) ? val.ToString() : val.ToString(format, culture));
      AddReplacementConverter<DateTime>((val, format) => string.IsNullOrWhiteSpace(format) ? val.ToString() : val.ToString(format, culture));
      AddReplacementConverter<int>((val, format) => string.IsNullOrWhiteSpace(format) ? val.ToString() : val.ToString(format, culture));
      AddReplacementConverter<uint>((val, format) => string.IsNullOrWhiteSpace(format) ? val.ToString() : val.ToString(format, culture));
      AddReplacementConverter<long>((val, format) => string.IsNullOrWhiteSpace(format) ? val.ToString() : val.ToString(format, culture));
      AddReplacementConverter<ulong>((val, format) => string.IsNullOrWhiteSpace(format) ? val.ToString() : val.ToString(format, culture));
      AddReplacementConverter<bool>((val, format) => string.IsNullOrWhiteSpace(format) ? val.ToString() : val.ToString(culture));
    }

    public CultureInfo Culture { get; set; } = new CultureInfo("en-EN");

    public string NoReplacementAvailableValue { get; set; } = "NoReplacementValueAvailable";

    public void AddReplacementConverter<TType>(Func<TType, string, string> converter)
    {
      this.converters[typeof(TType)] = (obj, ci,format) => converter((TType)obj, format);
    }

    public void AddFormatter<TType>(Func<TType, CultureInfo, string, string> converter)
    {
      this.converters[typeof(TType)] = (obj, ci, format) => converter((TType)obj, ci, format);
    }

    public void ReplaceTemplate(string pathDocx, IDictionary<string, object> replacments)
    {
      using (WordprocessingDocument doc = WordprocessingDocument.Open(pathDocx, true))
      {
        var body = doc.MainDocumentPart.Document.Body;
        var paragraphs = doc.MainDocumentPart.RootElement.Descendants<Paragraph>();
        foreach (var para in paragraphs)
        {
          var text = para.InnerText;
          var match = regex.IsMatch(text);
          if (match)
          {
            MergeParam(para);
            ReplaceTemplate(para, replacments);
          }
        }
        doc.Save();
      }
    }

    private void ReplaceTemplate(Paragraph para, IDictionary<string, object> replacments)
    {
      foreach (var textWithPlaceholder in para.Descendants<Text>().Where(t => regex.IsMatch(t.InnerText)))
      {
        var text = textWithPlaceholder.Text;
        textWithPlaceholder.Text = regex.Replace(text, new MatchEvaluator(match =>
        {
          var field = match.Groups["name"].Value;
          var format = match.Groups["format"]?.Value;
          if (replacments.TryGetValue(field, out var value))
          {
            if (converters.TryGetValue(value?.GetType(), out var converter))
            {
              return converter(value, Culture, format);
            }
            else
            {
              return value?.ToString() ?? "";
            }
          }
          return NoReplacementAvailableValue;
        }));
      }
    }
    private void MergeParam(Paragraph para)
    {
      var children = para.ChildElements.ToArray();
      var countOpen = 0;
      var countClose = 0;
      var elementsToMerge = new List<OpenXmlElement>();
      foreach (var child in children)
      {
        foreach (var c in child.InnerText)
        {
          if (c == '<') countOpen++;
          if (c == '>') countClose++;
        }
        if (countOpen > 0)
        {
          elementsToMerge.Add(child);
        }
        if (countOpen == countClose)
        {
          countOpen = 0;
          countClose = 0;
          Merge(para, elementsToMerge);
          elementsToMerge.Clear();
        }
      }
      Merge(para, elementsToMerge);
    }
    private void Merge(Paragraph para, List<OpenXmlElement> elementsToMerge)
    {
      if (elementsToMerge.Count == 0)
      {
        return;
      }
      var text = string.Join("", elementsToMerge.Select(e => e.InnerText));
      foreach (var element in elementsToMerge.ToArray())
      {
        var textElement = element.Descendants<Text>().FirstOrDefault();
        if (textElement != null)
        {
          textElement.Text = text;
          foreach (var elementToDelete in elementsToMerge.Where(e => e != element))
          {
            para.RemoveChild(elementToDelete);
          }
          return;
        }
      }

    }
  }
}
