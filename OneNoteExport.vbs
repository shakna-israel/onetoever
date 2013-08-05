Option Explicit

Const FILE_PATH = "C:\ExportFolder"


Sub PublishFirstPageOfFirstSectionOfFirstNotebookToWord()
    ' Connect to OneNote 2010.
    Dim oneNote As OneNote14.Application
    Set oneNote = New OneNote14.Application
    
    ' Get all of the Notebook nodes.
    Dim nodes As MSXML2.IXMLDOMNodeList
    Set nodes = GetFirstOneNoteNotebookNodes(oneNote)
    If Not nodes Is Nothing Then
        ' Get the first OneNote Notebook in the XML document.
        Dim node As MSXML2.IXMLDOMNode
        Set node = nodes(0)
        Dim noteBookName As String
        noteBookName = node.Attributes.getNamedItem("name").Text
        
        ' Get the ID for the Notebook so the code can retrieve
        ' the list of sections.
        Dim notebookID As String
        notebookID = node.Attributes.getNamedItem("ID").Text
               
        ' Load the XML for the Sections for the Notebook requested.
        Dim sectionsXml As String
        oneNote.GetHierarchy notebookID, hsSections, sectionsXml, xs2010
        
        Dim secDoc As MSXML2.DOMDocument
        Set secDoc = New MSXML2.DOMDocument
    
        If secDoc.LoadXML(sectionsXml) Then
            Dim secNodes As MSXML2.IXMLDOMNodeList
            Set secNodes = secDoc.DocumentElement.SelectNodes("//one:Section")

            If Not secNodes Is Nothing Then
                Dim secNode As MSXML2.IXMLDOMNode
                Set secNode = secNodes(0)
                
                Dim sectionID As String
                
                For Each secNode In secNodes
                                               
                sectionID = GetAttributeValueFromNode(secNode, "ID")
                                                      
                ' Load the XML for the Pages for the Section requested.
                Dim pagesXml As String
                oneNote.GetHierarchy sectionID, hsPages, pagesXml, xs2010
                
                Dim pagesDoc As MSXML2.DOMDocument
                Set pagesDoc = New MSXML2.DOMDocument
                
                If pagesDoc.LoadXML(pagesXml) Then
                    Dim pageNodes As MSXML2.IXMLDOMNodeList
                    Set pageNodes = pagesDoc.DocumentElement.SelectNodes("//one:Page")
                    
                    If Not pageNodes Is Nothing Then
                        Dim pageNode As MSXML2.IXMLDOMNode
                        Set pageNode = pageNodes(0)

                        Dim pageName As String
                        Dim pageID As String
                        
                        For Each pageNode In pageNodes
                            On Error Resume Next
                            pageName = GetAttributeValueFromNode(pageNode, "name")
                            pageID = GetAttributeValueFromNode(pageNode, "ID")
                            
                            ' Creating folder path for output for section
                            Dim sectionName As String
                            sectionName = GetAttributeValueFromNode(secNode, "name")
                            Dim sectionPath As String
                            sectionPath = sectionName & "\\"

                            ' Get the user's specified output folder.
                            Dim outputFolder As String
                            outputFolder = FILE_PATH
                        
                            ' Create a file name using the page's name.
                            Dim fileName As String
                            fileName = pageName & ".docx"
                        
                            ' Combine the two values into a single
                            ' Variable so it's easier to use twice.
                            Dim publishContentTo As String
                            publishContentTo = outputFolder & sectionPath & fileName
                        
                            ' Publish the page content to a PDF file
                            ' in the user specified folder.
                            ' The last parameter is not necessary unless you use the
                            ' pfEmf enum.
                            ' Note this method will fail if the file already exists
                            ' at the location specified.
                            oneNote.Publish pageID, publishContentTo, pfWord
                        Next

                    Else
                        MsgBox "OneNote 2010 Page nodes not found."
                    End If
                Else
                    MsgBox "OneNote 2010 Pages XML data failed to load."
                End If
                Next
            Else
                MsgBox "OneNote 2010 Section nodes not found."
            End If
        Else
            MsgBox "OneNote 2010 Section XML data failed to load."
        End If
    Else
        MsgBox "OneNote 2010 XML data failed to load."
    End If
    
End Sub

Private Function GetAttributeValueFromNode(node As MSXML2.IXMLDOMNode, attributeName As String) As String
    If node.Attributes.getNamedItem(attributeName) Is Nothing Then
        GetAttributeValueFromNode = "Not found."
    Else
        GetAttributeValueFromNode = node.Attributes.getNamedItem(attributeName).Text
    End If
End Function

Private Function GetFirstOneNoteNotebookNodes(oneNote As OneNote14.Application) As MSXML2.IXMLDOMNodeList
    ' Get the XML that represents the OneNote notebooks available.
    Dim notebookXml As String
    ' OneNote fills notebookXml with an XML document providing information
    ' about what OneNote notebooks are available.
    ' You want all the data and thus are providing an empty string
    ' for the bstrStartNodeID parameter.
    oneNote.GetHierarchy "", hsNotebooks, notebookXml, xs2010
    
    ' Use the MSXML Library to parse the XML.
    Dim doc As MSXML2.DOMDocument
    Set doc = New MSXML2.DOMDocument
    
    If doc.LoadXML(notebookXml) Then
        Set GetFirstOneNoteNotebookNodes = doc.DocumentElement.SelectNodes("//one:Notebook")
    Else
        Set GetFirstOneNoteNotebookNodes = Nothing
    End If
End Function