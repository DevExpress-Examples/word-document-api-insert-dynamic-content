using System;
using System.Diagnostics;
using System.Drawing;
using System.ServiceModel.Syndication;
using System.Xml;
using DevExpress.Office;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

namespace word_processing_insert_dynamic_content
{
    class Program
    {
        static void Main(string[] args)
        {
            SetupSecurityProtocol();

            RichEditDocumentServer wordProcessor = new RichEditDocumentServer();
            wordProcessor.LoadDocument("Dynamic Content.docx");
            InsertDocVariableField(wordProcessor.Document);
            wordProcessor.CalculateDocumentVariable += WordProcessor_CalculateDocumentVariable;
            wordProcessor.Document.Fields.Update();
            wordProcessor.SaveDocument("MergedDocument.docx", DocumentFormat.OpenXml);
            Process.Start(new ProcessStartInfo("MergedDocument.docx") { UseShellExecute = true });
        }

        private static void WordProcessor_CalculateDocumentVariable(object sender, CalculateDocumentVariableEventArgs e)
        {
            if (e.VariableName == "rssFeed")
            {
                e.KeepLastParagraph = true;
                e.Value = GenerateRssFeed();
                if (e.Value != null)
                    e.Handled = true;
                e.FieldLocked = true;
            }

        }

        private static void InsertDocVariableField(Document document)
        {
            document.Fields.Create(document.Range.End, "DOCVARIABLE rssFeed");
        }

        private static RichEditDocumentServer GenerateRssFeed()
        {
            RichEditDocumentServer rssProcessor = new RichEditDocumentServer();
            Document document = rssProcessor.Document;
            AbstractNumberingList abstractNumberingList = document.AbstractNumberingLists.BulletedListTemplate.CreateNew();
            document.NumberingLists.CreateNew(abstractNumberingList.Index);

            SyndicationFeed feed = null;
            try
            {
                using (XmlReader reader = XmlReader.Create("https://community.devexpress.com/blogs/MainFeed.aspx"))
                {
                    feed = SyndicationFeed.Load(reader);
                }
            }
            catch
            {
                return null;
            }
            document.BeginUpdate();
            foreach (SyndicationItem item in feed.Items)
                AddSyndicationItem(document, item);
            document.EndUpdate();
            return rssProcessor;
        }
        static void AddSyndicationItem(Document document, SyndicationItem item)
        {
            Paragraph paragraph = document.Paragraphs.Append();
            paragraph.LineSpacing = 1f;
            paragraph.ListIndex = 0;
            paragraph.SpacingAfter = 3;

            DocumentRange range = document.InsertText(paragraph.Range.Start, item.Title.Text);
            CharacterProperties properties = document.BeginUpdateCharacters(range);
            properties.FontSize = 12f;
            properties.FontName = "Segoe UI";
            document.EndUpdateCharacters(properties);

            if (item.Links.Count > 0)
            {
                Hyperlink hyperlink = document.Hyperlinks.Create(range);
                hyperlink.NavigateUri = item.Links[0].Uri.ToString();
            }

            range = document.InsertText(range.End, String.Format("{0}Published {1}", Characters.LineBreak, item.PublishDate.DateTime));
            properties = document.BeginUpdateCharacters(range);
            properties.FontSize = 8f;
            properties.FontName = "Segoe UI";
            properties.ForeColor = Color.Gray;
            document.EndUpdateCharacters(properties);
        }

        static void SetupSecurityProtocol()
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls11 | System.Net.SecurityProtocolType.Tls12;
            }
            catch
            {
            }
        }

    }
}
