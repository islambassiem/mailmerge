Attribute VB_Name = "Module1"

Sub DocAndPdfMailMergeDoLoop()                                                        ' name the subroutine "DocAndPdfMailMerge"
' Revised 18/11/2022

    Dim MasterDoc As Document, SingleMergeDoc As Document, LastRecordNum As Integer   ' Declare variables to store the Master Doc, the single merged file and the number of the last record to be merged
    Set MasterDoc = ActiveDocument                                               ' The master doc is the one with the merge fields as this is the where you will run the macro from so it is the active doc in the front

    MasterDoc.MailMerge.DataSource.ActiveRecord = wdLastRecord                   ' From the edit recipients list in MailMerge, go to the last record. This is to find where the end of the records are
    LastRecordNum = MasterDoc.MailMerge.DataSource.ActiveRecord                  ' Store the number of the position of the final record in the variable LastRecordNum so we know when to stop
    MasterDoc.MailMerge.DataSource.ActiveRecord = wdFirstRecord                  ' From the edit recipients list in MailMerge, go to the first record. Initiate the process to start from the beggining

    Do While LastRecordNum > 0                                                   ' we create this loop so that we can terminate the loop when we want. When we have completed the last record we set LastRecordNum to 0 which will terminate the loop

        MasterDoc.MailMerge.Destination = wdSendToNewDocument                                                 ' This sets the type of mail merge is to merge to document and not merge to email
        MasterDoc.MailMerge.DataSource.FirstRecord = MasterDoc.MailMerge.DataSource.ActiveRecord              ' We want to save only one merged documetnt at a time so we first want to merge from 1 to 1 then 2 to 2 ...
        MasterDoc.MailMerge.DataSource.LastRecord = MasterDoc.MailMerge.DataSource.ActiveRecord               ' ... . we therefore set the FirstRecord and LastRecord to the ActiveRecord. So we only merge the active record.
        MasterDoc.MailMerge.Execute False                                                                     ' This runs the MailMerge based on the FirstRecord and LastRecord range, so 1 record

        Set SingleMergeDoc = ActiveDocument                                      ' As the MailMerge has started, the merged document is now displayed infront so the merged document is now the active document as it is in the front, it is srored in variable "SingleMergeDoc"

        SingleMergeDoc.SaveAs2 _
            FileName:=MasterDoc.MailMerge.DataSource.DataFields("DocFolder").Value & Application.PathSeparator & _
                MasterDoc.MailMerge.DataSource.DataFields("FileName").Value & ".docx", _
            FileFormat:=wdFormatXMLDocument                                      ' The command SaveAs2 is the command to save a word doc and it will save the contents of "SingleMergeDoc" as a word docx in the folder found in path described in the field DocFolder in the data source, with the name as found in the field FileName in the data source

        SingleMergeDoc.ExportAsFixedFormat _
            OutputFileName:=MasterDoc.MailMerge.DataSource.DataFields("PdfFolder").Value & Application.PathSeparator & _
                MasterDoc.MailMerge.DataSource.DataFields("FileName").Value & ".pdf", _
            ExportFormat:=wdExportFormatPDF                                      ' The command ExportAsFixedFormat is used to save in format PDF and it will save the contents of "SingleMergeDoc" as a PDF documet in the folder found in path described in the field PdfFolder in the data source, with the name as found in the field FileName in the data source

        SingleMergeDoc.Close False                                                    ' Close the SingleMergeDoc to allow the variable SingleMergeDoc to accept another value which wil be the next merged document

        If MasterDoc.MailMerge.DataSource.ActiveRecord >= LastRecordNum Then     ' Check if the merged document we created was for the last record
            LastRecordNum = 0                                                    ' If we creaed the merged document for the last record tehn we set variable of LastRecordNum to zero so as to terminate the Do While loop
        Else
            MasterDoc.MailMerge.DataSource.ActiveRecord = wdNextRecord           ' If the documet we just merged was not from the last record then we move to the next record
        End If

    Loop                                                                         ' Jump back to the start of the DO While loop which will either terminate the loop if the LastRecordNum is 0 or repeat the process

End Sub

