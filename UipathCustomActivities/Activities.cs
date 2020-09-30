using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.IO;
using iTextSharp.text.pdf;

namespace UiPathCustomActivity
{
    public class PdfInvalidLink : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> FileName { get; set; }
        [Category("Output")]
        public OutArgument<List<string>> InvalidLinks { get; set; }
        private static List<string> GetPdfLinks(string file, int page)
        {
            List<string> Ret = new List<string>();
            using (PdfReader R = new PdfReader(file))
            {
                PdfDictionary PageDictionary = R.GetPageN(page);
                PdfArray Annots = PageDictionary.GetAsArray(PdfName.ANNOTS);

                if ((Annots == null) || (Annots.Length == 0))
                    return null;

                foreach (PdfObject A in Annots.ArrayList)
                {
                    PdfDictionary AnnotationDictionary = (PdfDictionary)PdfReader.GetPdfObject(A);

                    if (!AnnotationDictionary.Get(PdfName.SUBTYPE).Equals(PdfName.LINK))
                        continue;

                    if (AnnotationDictionary.Get(PdfName.A) == null)
                        continue;

                    var annotActionObject = AnnotationDictionary.Get(PdfName.A);
                    var AnnotationAction = (PdfDictionary)(annotActionObject.IsIndirect() ? PdfReader.GetPdfObject(annotActionObject) : annotActionObject);

                    if (AnnotationAction.Get(PdfName.S).Equals(PdfName.URI))
                    {
                        PdfString Destination = AnnotationAction.GetAsString(PdfName.URI);
                        if (Destination != null)
                            Ret.Add(Destination.ToString());
                    }
                }
                R.Close();
            }
            return Ret;
        }

        private static List<string> GetPdfInvalidLinks(List<string> links, int pageNo)
        {
            List<string> invalidLinks = new List<string>();
            if (links != null)
            {
                foreach (var item in links)
                {
                    if (!Uri.IsWellFormedUriString(item, UriKind.RelativeOrAbsolute))
                    {
                        invalidLinks.Add("URL: " + item);
                        invalidLinks.Add("Page: " + pageNo.ToString());
                    }
                }
            }
            return invalidLinks;
        }
        private static List<string> GetInvalidLinksFromPdf(string file)
        {
            List<string> invalid = new List<string>();
            using (PdfReader pdfReader = new PdfReader(file))
            {
                int pages = pdfReader.NumberOfPages;
                pdfReader.Close();
                for (int i = 1; i <= pages; i++)
                {
                    invalid = GetPdfInvalidLinks(GetPdfLinks(file, i), i);
                }
            }
            return invalid;
        }
        protected override void Execute(CodeActivityContext context)
        {
            string file = FileName.Get(context);
            InvalidLinks.Set(context, GetInvalidLinksFromPdf(file));
        }
    }

    public class ReplaceXmlEntity : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> XmlFileName { get; set; }
        private static void replaceXml(string fileName)
        {
            string textOutput = File.ReadAllText(fileName);
            String match = "(?:(Alt|src)=\\\")([\\s\\S]*?)\"([.*\">]|[.*\"/>])";
            String ActualMatch = "(?:(ActualText)=\\\")([\\s\\S]*?)\"";
            Regex re = new Regex(match);
            MatchCollection mcl = re.Matches(textOutput);
            string newTextt = "";
            if (mcl.Count == 0)
            {
                newTextt = textOutput;
            }
            else
            {
                foreach (Match m in mcl)
                {
                    if (m.ToString().Contains("src"))
                    {
                        newTextt = textOutput.Replace(m.Value, string.Empty + "/");
                    }
                    else
                        newTextt = textOutput.Replace(m.Value, "Alt = \"Rpa Pdf\"" + ">");
                    textOutput = newTextt;
                }
            }
            Regex re1 = new Regex(ActualMatch);
            MatchCollection mcl2 = re1.Matches(textOutput);
            newTextt = "";
            if (mcl2.Count == 0)
            {
                newTextt = textOutput;
            }
            else
            {
                foreach (Match m in mcl2)
                {
                    newTextt = textOutput.Replace(m.Value, m.Value.Replace("&", "&amp;").Replace(">", "&gt;").Replace("<", "&lt;"));
                    textOutput = newTextt;
                }
            }
            File.WriteAllText(fileName, textOutput);
        }
        protected override void Execute(CodeActivityContext context)
        {
            string fileName = XmlFileName.Get(context);
            replaceXml(fileName);
        }
    }
}
