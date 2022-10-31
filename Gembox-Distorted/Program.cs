// See https://aka.ms/new-console-template for more information
using GemBox.Document;
ComponentInfo.SetLicense("LICENSE");

Console.WriteLine($"GemBox version: {ComponentInfo.FullVersion}");

var document = DocumentModel.Load(@"C:\Gembox\Distorted.doc", new DocLoadOptions { PreserveUnsupportedFeatures = true });
bool hasUnsupportedElement = document.GetChildElements(true, ElementType.PreservedInline, ElementType.PreservedDrawingElement).Any();
Console.Write($"Har ikke understøttede elementer: {hasUnsupportedElement}");

if (hasUnsupportedElement)
{
    foreach (var a in document.GetChildElements(true, ElementType.PreservedInline, ElementType.PreservedDrawingElement))
    {
        Console.WriteLine(a);
    }
}

MemoryStream ms = new MemoryStream();
document.Save(ms, GemBox.Document.SaveOptions.DocxDefault);

using (FileStream fs = new FileStream(@"C:\Gembox\Distorted.docx", FileMode.Create))
{
    ms.Position = 0;
    ms.CopyTo(fs);
}