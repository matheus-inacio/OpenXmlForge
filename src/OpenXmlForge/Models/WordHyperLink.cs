using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlForge.Models;

public sealed class WordHyperLink
{
    public string Url { get; set; } = string.Empty;

    public string Text { get; set; } = string.Empty;

    public UnderlineValues UnderLineValue { get; set; } = UnderlineValues.Single;

    public TargetFrameType TargetFrame { get; set; } = TargetFrameType.Blank;

    internal string GetTargetFrameString() => TargetFrame switch
    {
        TargetFrameType.Blank => "_blank",
        TargetFrameType.Top => "_top",
        TargetFrameType.Self => "_self",
        TargetFrameType.Parent => "_parent",
        _ => "_blank",
    };
}
