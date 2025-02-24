Imports System.Drawing
Imports System.ServiceModel.Syndication
Imports System.Xml
Imports DevExpress.Office
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native

Module Module1

    Sub Main(ByVal args() As String)
        SetupSecurityProtocol()

        Dim wordProcessor As New RichEditDocumentServer()
        wordProcessor.LoadDocument("Dynamic Content.docx")
        InsertDocVariableField(wordProcessor.Document)
        AddHandler wordProcessor.CalculateDocumentVariable, AddressOf WordProcessor_CalculateDocumentVariable
        wordProcessor.Document.Fields.Update()
        wordProcessor.SaveDocument("MergedDocument.docx", DocumentFormat.OpenXml)
        Process.Start(New ProcessStartInfo("MergedDocument.docx") With {.UseShellExecute = True})
    End Sub

    Private Sub WordProcessor_CalculateDocumentVariable(ByVal sender As Object, ByVal e As CalculateDocumentVariableEventArgs)
        If e.VariableName = "rssFeed" Then
            e.KeepLastParagraph = True
            e.Value = GenerateRssFeed()
            If e.Value IsNot Nothing Then
                e.Handled = True
            End If
            e.FieldLocked = True
        End If

    End Sub

    Private Sub InsertDocVariableField(ByVal document As Document)
        document.Fields.Create(document.Range.End, "DOCVARIABLE rssFeed")
    End Sub

    Private Function GenerateRssFeed() As RichEditDocumentServer
        Dim rssProcessor As New RichEditDocumentServer()
        Dim document As Document = rssProcessor.Document
        Dim abstractNumberingList As AbstractNumberingList = document.AbstractNumberingLists.BulletedListTemplate.CreateNew()
        document.NumberingLists.CreateNew(abstractNumberingList.Index)

        Dim feed As SyndicationFeed = Nothing
        Try
            Using reader As XmlReader = XmlReader.Create("https://community.devexpress.com/blogs/MainFeed.aspx")
                feed = SyndicationFeed.Load(reader)
            End Using
        Catch
            Return Nothing
        End Try
        document.BeginUpdate()
        For Each item As SyndicationItem In feed.Items
            AddSyndicationItem(document, item)
        Next item
        document.EndUpdate()
        Return rssProcessor
    End Function
    Private Sub AddSyndicationItem(ByVal document As Document, ByVal item As SyndicationItem)
        Dim paragraph As Paragraph = document.Paragraphs.Append()
        paragraph.LineSpacing = 1.0F
        paragraph.ListIndex = 0
        paragraph.SpacingAfter = 3

        Dim range As DocumentRange = document.InsertText(paragraph.Range.Start, item.Title.Text)
        Dim properties As CharacterProperties = document.BeginUpdateCharacters(range)
        properties.FontSize = 12.0F
        properties.FontName = "Segoe UI"
        document.EndUpdateCharacters(properties)

        If item.Links.Count > 0 Then
            Dim hyperlink As Hyperlink = document.Hyperlinks.Create(range)
            hyperlink.NavigateUri = item.Links(0).Uri.ToString()
        End If

        range = document.InsertText(range.End, String.Format("{0}Published {1}", Characters.LineBreak, item.PublishDate.DateTime))
        properties = document.BeginUpdateCharacters(range)
        properties.FontSize = 8.0F
        properties.FontName = "Segoe UI"
        properties.ForeColor = Color.Gray
        document.EndUpdateCharacters(properties)
    End Sub

    Private Sub SetupSecurityProtocol()
        Try
            System.Net.ServicePointManager.SecurityProtocol = System.Net.ServicePointManager.SecurityProtocol Or System.Net.SecurityProtocolType.Tls11 Or System.Net.SecurityProtocolType.Tls12
        Catch
        End Try
    End Sub


End Module
