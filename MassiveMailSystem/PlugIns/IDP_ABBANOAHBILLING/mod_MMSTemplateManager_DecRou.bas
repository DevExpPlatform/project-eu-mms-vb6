Attribute VB_Name = "mod_MMSTemplateManager_DecRou"
Option Explicit

Private Type strct_TemplateInfo
    ID_DESCR                As String
    FIELDS                  As String
    FIELDSDATA              As String
    FIELDSNUM               As Long
    TEMP_BITWISE            As Double
    TEMP_ID                 As Long
    TEMP_PAGES              As Long
End Type

Private Type strct_SectionInfo
    SECTIONDESCRIPTION      As String
    TEMPLATES()             As strct_TemplateInfo
End Type

Private Type strct_DynamicDataSupport
    DATAKEY                 As String
    DATAVALUE               As String
End Type

Private WS_CNTRFIELD        As Integer
Private WS_DDSCNTR          As Integer
Private WS_DDS()            As strct_DynamicDataSupport

Public WS_SECTIONS()        As strct_SectionInfo
Public WS_FLG_SCTN_ATTACH   As Boolean

Public Sub DDS_ADD(WS_DATA_KEY As String, WS_DATA_VALUE As String)

    WS_DDSCNTR = (WS_DDSCNTR + 1)
    ReDim Preserve WS_DDS(WS_DDSCNTR) As strct_DynamicDataSupport

    WS_DDS(WS_DDSCNTR).DATAKEY = WS_DATA_KEY
    WS_DDS(WS_DDSCNTR).DATAVALUE = WS_DATA_VALUE

End Sub

Public Sub DDS_INIT()

    Erase WS_DDS
    
    WS_CNTRFIELD = 0
    WS_DDSCNTR = -1

End Sub

Public Function GET_FIELDS(WS_FIELD As String, WS_PAGESNUM As Integer) As String
    
    Dim objFieldsDoc            As MSXML2.DOMDocument60
    Dim objFieldNode            As IXMLDOMNode
    Dim objFieldNodeList        As IXMLDOMNodeList
    Dim objPropertiesNode       As IXMLDOMNode
    Dim objPropertiesNodeList   As IXMLDOMNodeList
    
    Set objFieldsDoc = New MSXML2.DOMDocument60
    objFieldsDoc.async = False
    
    If (objFieldsDoc.loadXML("<fields>" & WS_FIELD & "</fields>")) Then
        Set objFieldNodeList = objFieldsDoc.selectNodes("/fields/field")
        
        For Each objFieldNode In objFieldNodeList
            If (objFieldNode.Attributes.getNamedItem("id").nodeValue = "TXT_ANNEXED_PXXX") Then
                WS_CNTRFIELD = (WS_CNTRFIELD + 1)
                objFieldNode.Attributes.getNamedItem("id").nodeValue = "TXT_ANNEXED_P" & Format$(WS_CNTRFIELD, "000")
            End If
            
            Set objPropertiesNodeList = objFieldNode.selectNodes("properties")
            
            For Each objPropertiesNode In objPropertiesNodeList
                objPropertiesNode.Attributes.getNamedItem("pageid").nodeValue = (objPropertiesNode.Attributes.getNamedItem("pageid").nodeValue + WS_PAGESNUM)
            Next objPropertiesNode
            
            GET_FIELDS = GET_FIELDS & objFieldNode.xml
        Next objFieldNode
        
        Set objFieldNode = Nothing
        Set objFieldNodeList = Nothing
        Set objPropertiesNode = Nothing
        Set objPropertiesNodeList = Nothing
    End If

    Set objFieldsDoc = Nothing

End Function

Public Function GET_FIELDSDATA(ByVal WS_FIELDSDATA As String) As String
    
    Dim I As Integer
    
    For I = 0 To UBound(WS_DDS)
        WS_FIELDSDATA = Replace$(WS_FIELDSDATA, WS_DDS(I).DATAKEY, WS_DDS(I).DATAVALUE)
    Next I

    WS_CNTRFIELD = 0
    GET_FIELDSDATA = Left$(WS_FIELDSDATA, Len(WS_FIELDSDATA) - 4)
    
End Function

Public Sub TEMPLATES_MANAGER_INIT(XMLSrc As String)

    Dim I                       As Integer
    Dim J                       As Integer
    Dim objDoc                  As MSXML2.DOMDocument60
    Dim objFieldNodeList        As IXMLDOMNodeList
    Dim objFieldsDoc            As MSXML2.DOMDocument60
    Dim objFieldValueNodelist   As IXMLDOMNodeList
    Dim objSectionNode          As IXMLDOMNode
    Dim objSectionNodelist      As IXMLDOMNodeList
    Dim objTemplateNode         As IXMLDOMNode
    Dim objTemplateNodelist     As IXMLDOMNodeList
    Dim WS_CNTRBITWISE          As Integer
    Dim WS_CNTRSCTN             As Integer
    Dim WS_CNTRTMPLT            As Integer
    Dim WS_XML                  As String
    
    Set objDoc = New MSXML2.DOMDocument60
    objDoc.async = False
    
    WS_CNTRBITWISE = -1
    WS_CNTRFIELD = 0
    WS_CNTRSCTN = -1
    WS_FLG_SCTN_ATTACH = False
    
    Erase WS_SECTIONS
    
    If (objDoc.loadXML(XMLSrc)) Then
        Set objSectionNodelist = objDoc.selectNodes("//section")
                
        For Each objSectionNode In objSectionNodelist
            WS_CNTRSCTN = (WS_CNTRSCTN + 1)
            WS_CNTRTMPLT = -1
            
            ReDim Preserve WS_SECTIONS(WS_CNTRSCTN) As strct_SectionInfo
                    
            With WS_SECTIONS(WS_CNTRSCTN)
                .SECTIONDESCRIPTION = objSectionNode.Attributes.getNamedItem("description").Text
                
                If (.SECTIONDESCRIPTION = "ATTACHMENT") Then WS_FLG_SCTN_ATTACH = True
                
                Set objTemplateNodelist = objSectionNode.selectNodes("template")
            
                For Each objTemplateNode In objTemplateNodelist
                    WS_CNTRTMPLT = (WS_CNTRTMPLT + 1)
                    WS_CNTRBITWISE = (WS_CNTRBITWISE + 1)
                    
                    ReDim Preserve .TEMPLATES(WS_CNTRTMPLT) As strct_TemplateInfo

                    With .TEMPLATES(WS_CNTRTMPLT)
                        .FIELDSNUM = -1
                        .TEMP_ID = IIf((objTemplateNode.Attributes.getNamedItem("id").Text = ""), -1, objTemplateNode.Attributes.getNamedItem("id").Text)
                        .TEMP_PAGES = IIf((objTemplateNode.Attributes.getNamedItem("pages").Text = ""), -1, objTemplateNode.Attributes.getNamedItem("pages").Text)
                        .TEMP_BITWISE = (2 ^ WS_CNTRBITWISE)
                        
                        If (Not objTemplateNode.Attributes.getNamedItem("datafile") Is Nothing) Then
                            Set objFieldsDoc = New MSXML2.DOMDocument60
                            objFieldsDoc.async = False
                            
                            WS_XML = GET_EXTERNALINFO(DLLParams.EXTRASPATH & objTemplateNode.Attributes.getNamedItem("datafile").Text)
                            
                            If (objFieldsDoc.loadXML(WS_XML)) Then
                                If (Not objFieldsDoc.getElementsByTagName("document").Item(0).Attributes.getNamedItem("id") Is Nothing) Then
                                    .ID_DESCR = objTemplateNode.Attributes.getNamedItem("datafile").Text
                                End If
                            
                                Set objFieldNodeList = objFieldsDoc.selectNodes("/document/fields/field")
                                Set objFieldValueNodelist = objFieldsDoc.selectNodes("/document/fieldsvalue/fieldvalue")
                                        
                                For I = 0 To (objFieldNodeList.length - 1)
                                    If (Trim$(objFieldNodeList.Item(I).Attributes.getNamedItem("id").nodeValue) = "") Then
                                        objFieldNodeList.Item(I).Attributes.getNamedItem("id").nodeValue = "TXT_ANNEXED_PXXX"
                                        
                                        .FIELDS = .FIELDS & objFieldNodeList.Item(I).xml
                                        
                                        If (objFieldValueNodelist.length > 0) Then
                                            For J = 0 To (objFieldValueNodelist.Item(I).childNodes.length - 1)
                                                If (Left$(objFieldValueNodelist.Item(I).childNodes.Item(J).xml, 4) <> "<!--") Then
                                                    .FIELDSDATA = .FIELDSDATA & objFieldValueNodelist.Item(I).childNodes.Item(J).xml
                                                End If
                                            Next J
                                            
                                            .FIELDSDATA = .FIELDSDATA & "[ED]"
                                            .FIELDSNUM = (.FIELDSNUM + 1)
                                        End If
                                    Else
                                        .FIELDS = .FIELDS & objFieldNodeList.Item(I).xml
                                    End If
                                Next I
                            
                                Set objFieldNodeList = Nothing
                                Set objFieldValueNodelist = Nothing
                            End If
                            
                            Set objFieldsDoc = Nothing
                        End If
                    End With
                Next objTemplateNode
            End With
        Next objSectionNode
    End If
    
    Set objSectionNode = Nothing
    Set objSectionNodelist = Nothing
    Set objTemplateNode = Nothing
    Set objTemplateNodelist = Nothing
    Set objDoc = Nothing

End Sub
