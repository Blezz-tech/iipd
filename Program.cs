using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System;
using System.Linq;
using System.Diagnostics; 


string filepath = @"./bin/source.docx";

var psi = new ProcessStartInfo {
    FileName = "nu",
    Arguments = "generate.nu",
    UseShellExecute = false,
    RedirectStandardOutput = true,
    RedirectStandardError = true
};

var process = new Process {
    StartInfo = psi
};

process.Start();

string output = process.StandardOutput.ReadToEnd();
string error = process.StandardError.ReadToEnd();

process.WaitForExit();


// using (WordprocessingDocument document = WordprocessingDocument.Open(filepath, true))
// {
//     MainDocumentPart? MainDocumentPart = document.MainDocumentPart;
//     StyleDefinitionsPart? styleDefinitionsPart = MainDocumentPart?.StyleDefinitionsPart;
//     Styles? styles = styleDefinitionsPart?.Styles;

//     // По умолчанию
//     RunPropertiesDefault rPrDef = styles?.DocDefaults?.RunPropertiesDefault;
//     RunPropertiesBaseStyle rPrDefrPr = rPrDef?.RunPropertiesBaseStyle;

//     ParagraphPropertiesDefault  pPrDef = styles?.DocDefaults?.ParagraphPropertiesDefault;
//     ParagraphPropertiesBaseStyle pPrDefpPr = pPrDef?.ParagraphPropertiesBaseStyle;

//     rPrDefrPr.RunFonts = new RunFonts() {
//         HighAnsi = "Times New Roman",
//         Ascii = "Times New Roman"
//     };
//     rPrDefrPr.FontSize = new FontSize() { Val = "14pt" };
//     rPrDefrPr.FontSizeComplexScript = new FontSizeComplexScript() { Val = "14pt" };
//     rPrDefrPr.Color = new Color() { Val = "000000" };
//     rPrDefrPr.Languages = new Languages() {
//         Val = "ru-RU",
//         Bidi = "ru-RU",
//         EastAsia = "ru-RU"
//     };

//     pPrDefpPr.SpacingBetweenLines = new SpacingBetweenLines () {
//         Line = "360",
//         After = "0",
//         Before = "0"
//     };

//     foreach (Style style in styles.Elements<Style>()) {
//         string? styleValue = style?.StyleId;

//         StyleRunProperties rPr = style.Elements<StyleRunProperties>().FirstOrDefault();
//         StyleParagraphProperties pPr = style.Elements<StyleParagraphProperties>().FirstOrDefault();

//         if (rPr == null && style.Type == "paragraph") {
//             rPr = new StyleRunProperties();
//             style.Append(rPr);
//         }

//         if (pPr == null && style.Type == "paragraph") {
//             pPr = new StyleParagraphProperties();
//             style.Append(pPr);
//         }

//         switch (styleValue) {
//             case "Normal": {
//                 rPr.RunFonts = new RunFonts() {
//                     HighAnsi = "Times New Roman",
//                     Ascii = "Times New Roman"
//                 };
//                 rPr.FontSize = new FontSize() { Val = "14pt" };
//                 rPr.Color = new Color() { Val = "000000" };

//                 pPr.SpacingBetweenLines = new SpacingBetweenLines() {
//                     Line = "360",
//                     LineRule = LineSpacingRuleValues.Auto,
//                     After = "0",
//                     Before = "0"
//                 };
//                 pPr.Indentation = new Indentation() {
//                     FirstLine = "709",
//                     Left = "0",
//                     Right = "0"
//                 };
//                 pPr.Justification = new Justification() {
//                     Val = JustificationValues.Both
//                 };
//                 break;
//             }
            
//             case "BodyText": {
//                 pPr.SpacingBetweenLines = new SpacingBetweenLines() {
//                     After = "0",
//                     Before = "0"
//                 };
//                 break;
//             }

//             case "Heading1": {
//                 rPr.RunFonts = new RunFonts() {
//                     HighAnsi = "Times New Roman",
//                     Ascii = "Times New Roman"
//                 };
//                 rPr.FontSize = new FontSize() { Val = "16pt" };
//                 rPr.Color = new Color() { Val = "000000" };

//                 pPr.SpacingBetweenLines = new SpacingBetweenLines() {
//                     Line = "360",
//                     LineRule = LineSpacingRuleValues.Auto,
//                     After = "0",
//                     Before = "0"
//                 };
//                 pPr.Indentation = new Indentation() {
//                     FirstLine = "709",
//                     Left = "0",
//                     Right = "0"
//                 };
//                 pPr.Justification = new Justification() {
//                     Val = JustificationValues.Left
//                 };

//                 break;
//             }
            
//             case "Compact": {
//                 pPr.SpacingBetweenLines = new SpacingBetweenLines() {
//                     After = "0",
//                     Before = "0"
//                 };

//                 break;
//             }

//             // Картинка
//             case "CaptionedFigure": {
//                 pPr.SpacingBetweenLines = new SpacingBetweenLines() {
//                     Line = "360",
//                     LineRule = LineSpacingRuleValues.Auto,
//                     After = "0",
//                     Before = "0"
//                 };
//                 pPr.Indentation = new Indentation() {
//                     FirstLine = "0",
//                     Left = "0",
//                     Right = "0"
//                 };
//                 pPr.Justification = new Justification() {
//                     Val = JustificationValues.Center
//                 };

//                 break;
//             }

//             // Текст под картинкой
//             case "ImageCaption": {
//                 rPr.RunFonts = new RunFonts() {
//                     HighAnsi = "Times New Roman",
//                     Ascii = "Times New Roman"
//                 };
//                 rPr.FontSize = new FontSize() { Val = "10pt" };
//                 rPr.Color = new Color() { Val = "000000" };
//                 rPr.Bold = new Bold() {Val = true};
//                 rPr.Italic = new Italic() {Val = false};

//                 pPr.SpacingBetweenLines = new SpacingBetweenLines() {
//                     Line = "360",
//                     LineRule = LineSpacingRuleValues.Auto,
//                     After = "0",
//                     Before = "0"
//                 };
//                 pPr.Indentation = new Indentation() {
//                     FirstLine = "0",
//                     Left = "0",
//                     Right = "0"
//                 };
//                 pPr.Justification = new Justification() {
//                     Val = JustificationValues.Center
//                 };


//                 break;
//             }
//         }
//     }
//     styleDefinitionsPart.Styles.Save();
// }





