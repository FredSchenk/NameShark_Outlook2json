Attribute VB_Name = "NameShark_Outlook2json"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'set a root folder for all the files/folders to be created. Change this to your needs...
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const ROOTFOLDER = "DRIVE:\PATH\TO\FOLDER\NameShark_Outlook2json\"


Sub NS_CreateNamesharkJSON()

Dim objOL As Outlook.Application
Dim objNS As Outlook.NameSpace
Dim objContactsFolder As Outlook.MAPIFolder
Dim objItems As Outlook.Items
Dim objContact As Outlook.ContactItem
Dim objAttachments As Attachments
Dim objAttachment As Attachment
Dim iDone As Integer
Dim iChanged As Integer
Dim sFilter As String
Dim sGroup, sFolder, sFile, sFirst, sMiddle, sLast, sGender, sDescription, sEnc, sNameShark As String
Dim bHasPicture As Boolean
Dim objXML As MSXML2.DOMDocument60
Dim objNode As MSXML2.IXMLDOMElement

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'examples of group names and search strings - just enter your own last
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sGroup = "Family / Friends"
    sFilter = "[Categories] = '" & sGroup & "'"
    sGroup = Replace(sGroup, "/", "-")  'no forward slash in the cat-name
    
    sGroup = "Workplace"
    sFilter = "@SQL=""urn:schemas:httpmail:textdescription"" ci_phrasematch 'MY_WORK_PLACE' OR ""urn:schemas:contacts:o"" ci_phrasematch 'MY_WORK_PLACE'"
    
    sGroup = "SomeText"
    sFilter = "@SQL=""urn:schemas:httpmail:textdescription"" LIKE '%Some Text%'"
    
    sGroup = "Some_Other_Text"
    sFilter = "@SQL=""urn:schemas:httpmail:textdescription"" ci_phrasematch 'Some (Other) Text'"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'examples of group names and search strings - just enter your own last
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    On Error Resume Next
    
    'set the outlook objects and specify the items in the default contacts folder using the filter
    Set objOL = CreateObject("Outlook.Application")
    Set objNS = objOL.GetNamespace("MAPI")
    Set objContactsFolder = objNS.GetDefaultFolder(olFolderContacts) 'Get default Contacts Folder, if you want another folder you'll need to edit this command
    Set objItems = objContactsFolder.Items.Restrict(sFilter) 'Only retrieve the contacts using the filter
    
    ' set the counters to zero and initialise the progresss window
    iDone = 0
    iChanged = 0
    Progress.Show
    Progress.tTotal = objItems.Count
    
    ' loop through all the found items
    For itemCounter = 1 To objItems.Count
    
        Set objContact = objItems(itemCounter)
        Set objAttachments = objContact.Attachments
        iDone = iDone + 1
        bHasPicture = False
    
        For Each objAttachment In objAttachments
          If objAttachment.FileName = "ContactPicture.jpg" Then 'only processs the picture of the contact, not any other attachments
            If (objContact.Gender = olMale) Then
                sGender = "male"
            ElseIf (objContact.Gender = olFemale) Then
                sGender = "female"
            Else
                Progress.UpdateCounters iDone, iChanged, "*** SKIP *** : " & objContact.FileAs ' skip any contacts without gender
                Exit For
            End If
            sFirst = Trim(objContact.FirstName)
            sMiddle = Trim(objContact.MiddleName)
            sLast = Trim(objContact.LastName)
            sDescription = Trim(objContact.JobTitle) & " @ " & Trim(objContact.CompanyName) ' set the job title and company name as description in NameShark. Other options are possible too...
            
            ' export the contact picture to the group-folder with the contacts name as filename...could have used the 'FileAs'-property but chose not to.
            sFolder = ROOTFOLDER & sGroup & "\"
            sFile = (sFirst & " " & Trim(sMiddle & " " & sLast))
            sExt = ".jpg"
            
            If (Dir(sFolder, vbDirectory) = "") Then
                MkDir (sFolder) ' the group-folder should exist after saving the first picture - this could be placed outside the loop
            End If
            
            objAttachment.SaveAsFile (sFolder & sFile & sExt) ' actually save the contact picture
            
            ' create a base64 string from the just saved contact picture and remove the newline statements
            Set objXML = New MSXML2.DOMDocument60
            Set objNode = objXML.createElement("b64")

            objNode.dataType = "bin.base64"
            objNode.nodeTypedValue = NS_ReadFile(sFolder & sFile & sExt)
            sEnc = objNode.Text
            
            Set objNode = Nothing
            Set objXML = Nothing
              
            sEnc = Replace(sEnc, Chr(10), vbNullString, , , vbTextCompare)
            
            ' create the NameShark-string using "NS_Safestring"-ed values of the names and desciption
            sNameShark = "{""first"":""" & NS_Safestring(sFirst) & ""","
            sNameShark = sNameShark & """last"":""" & NS_Safestring(Trim(sMiddle & " " & sLast)) & ""","
            sNameShark = sNameShark & """gender"":""" & sGender & ""","
            sNameShark = sNameShark & """details"":""" & NS_Safestring(sDescription) & ""","
            sNameShark = sNameShark & """photoData"":""data:image/jpeg;base64," & sEnc & """}"
                      
            ' write the NameShark-string to the NameShark-file, append it to any existing file contents
            Call NS_WriteStringToFile(sNameShark, ROOTFOLDER & "NameSharkGroup-" & sGroup & ".json", "{""name"":""" & NS_Safestring(sGroup) & """,""contacts"":[", True)
            
            ' Update the counter and show the progress
            iChanged = iChanged + 1
            Progress.UpdateCounters iDone, iChanged, sFile
          Else
            ' show the progress
            Progress.UpdateCounters iDone, iChanged, "*** NoImage : " & objContact.FileAs
          End If
        Next
    
    Next
    ' write the closing-statement to the file, do not append with a comma
    Call NS_WriteStringToFile("]}", ROOTFOLDER & "NameSharkGroup-" & sGroup & ".json", "", False)
    
    'Clean up all the set objects
    Set objAttachments = Nothing
    Set objContact = Nothing
    Set objItems = Nothing
    Set objContactsFolder = Nothing
    Set objNS = Nothing
    Set objOL = Nothing
    
    'Enable the close button on the progress form to indicate the process has ended
    Progress.btnClose.Enabled = True

End Sub


Private Function NS_ReadFile(sFilename)
Const adTypeBinary = 1          ' Binary file

Dim objStream

    ' Open data stream from file
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = adTypeBinary
    objStream.Open
    objStream.LoadFromFile (sFilename)

    NS_ReadFile = objStream.Read()

End Function

                            
Private Sub NS_WriteStringToFile(sString, sFilename, sStartString, bAppendWithComma)
' write sString to sFilename. If the file exists, start sString with a comma to seperate statements if bAppendWithComma is True. If the file does not exists, create the file and start it with sStartString before writing sString

Dim oFso As New FileSystemObject
Dim oFile As TextStream

    Set oFso = CreateObject("Scripting.FileSystemObject")
    If oFso.FileExists(sFilename) Then
        Set oFile = oFso.OpenTextFile(sFilename, ForAppending, False) ' append
        If bAppendWithComma Then sString = "," & sString ' put a comma between statements
    Else
        Set oFile = oFso.CreateTextFile(sFilename, True) ' create
        oFile.WriteLine sStartString
    End If
    
    oFile.WriteLine sString
    oFile.Close

    Set oFile = Nothing
    Set oFso = Nothing

End Sub


Private Function NS_Safestring(sText)
Const AccChars = "ŠŽšžŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ"
Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"

    For i = 1 To Len(AccChars)
        a = Mid(AccChars, i, 1)
        R = Mid(RegChars, i, 1)
        sText = Replace(sText, a, R)
    Next
    NS_Safestring = sText

End Function
