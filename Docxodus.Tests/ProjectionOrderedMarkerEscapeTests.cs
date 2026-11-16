using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A plain paragraph whose text starts with digits and a period ("3. Lorem", "2026. Lorem") is
/// literal text, but the projection emitted it unescaped. Echoing that line back through
/// <see cref="DocxSession.ReplaceText"/> — the round trip the projection promises an agent — parsed
/// it as an ordered-list item: the paragraph became a numbered list item and lost its number.
/// </summary>
public class ProjectionOrderedMarkerEscapeTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static byte[] BuildParagraphs(params OpenXmlElement[][] paragraphs)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            foreach (var runs in paragraphs)
                main.Document.Body!.Append(new Paragraph(new Run(runs)));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static OpenXmlElement[] Text(string text) =>
        new OpenXmlElement[] { new Text(text) { Space = SpaceProcessingModeValues.Preserve } };

    /// <summary>The projected text of the block with this anchor: everything after "{#id} ".</summary>
    private static string ProjectedText(DocxSession session, string anchorId)
    {
        var markdown = session.Project().Markdown;
        var start = markdown.IndexOf("{#" + anchorId + "} ", System.StringComparison.Ordinal);
        Assert.True(start >= 0, $"anchor {anchorId} not in projection");
        start += anchorId.Length + 4;
        var end = markdown.IndexOf("\n\n", start, System.StringComparison.Ordinal);
        return markdown[start..(end < 0 ? markdown.Length : end)];
    }

    private static string ParagraphText(DocxSession session, string anchorId) =>
        string.Concat(XElement.Parse(session.Raw.GetXml(anchorId)).Descendants(W + "t").Select(t => (string)t));

    [Theory]
    [InlineData("3. Lorem ipsum dolor sit amet.")]
    [InlineData("2026. Lorem ipsum.")]
    [InlineData("7.")]
    public void EchoingAProjectedParagraph_ThatStartsWithANumberAndPeriod_KeepsItAPlainParagraph(string text)
    {
        using var session = new DocxSession(BuildParagraphs(Text("Intro."), Text(text)));
        var anchor = session.Project().AnchorIndex.Keys.ElementAt(1);

        var echoed = session.ReplaceText(anchor, ProjectedText(session, anchor));

        Assert.True(echoed.Success, echoed.Error?.Message);
        var modified = Assert.Single(echoed.Modified!);
        Assert.Equal("p", modified.Kind);
        Assert.Null(XElement.Parse(session.Raw.GetXml(modified.Id)).Element(W + "pPr")?.Element(W + "numPr"));
        Assert.Equal(text, ParagraphText(session, modified.Id));
    }

    [Fact]
    public void Projection_EscapesTheMarkerPeriod_OnlyAtTheStartOfALine()
    {
        using var session = new DocxSession(BuildParagraphs(
            Text("3. Lorem"),
            Text("12.5 percent, 4. not a marker mid line"),
            new OpenXmlElement[] { new Text("First line"), new Break(), new Text("2. second line") }));
        var anchors = session.Project().AnchorIndex.Keys.ToList();

        Assert.Equal("3\\. Lorem", ProjectedText(session, anchors[0]));
        Assert.Equal("12.5 percent, 4. not a marker mid line", ProjectedText(session, anchors[1]));
        Assert.Equal("First line  \n2\\. second line", ProjectedText(session, anchors[2]));
    }
}
