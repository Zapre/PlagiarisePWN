Imports System.IO
Imports Microsoft.Office.Interop.Word
Imports System.Text.RegularExpressions
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If phraseTXT.Text = "" Then
                MsgBox("EMPTY!", MsgBoxStyle.Critical)
            Else
                scanspecific()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Function scanspecific()
        Dim quote As String = quoter.Text
        Try
            specificBROWSE.Navigate("https://www.google.com/search?q=" + quote + phraseTXT.Text + quote)
            bingbrowse.Navigate("http://bing.com/search?q=" + quote + phraseTXT.Text + quote)
            duckduckgobrowse.Navigate("https://duckduckgo.com/?q=" + quote + phraseTXT.Text + quote)
        Catch ex As Exception

        End Try
    End Function

    Function selectitem()
        Try
            specificURLS.SelectedItem = specificURLS.Items.Item(0)
        Catch ex As Exception

        End Try
    End Function

    Private Sub checksrc_ProgressChanged(sender As Object, e As WebBrowserProgressChangedEventArgs) Handles checksrc.ProgressChanged
        Try
            ProgressBar1.Value = CType(((e.CurrentProgress / e.MaximumProgress) * 100), Integer)
        Catch ex As Exception

        End Try
    End Sub

    Function removebing()
        Try
            For i As Integer = specificURLS.Items.Count - 1 To 0 Step -1
                If specificURLS.Items(i).Contains("bing.com") Then
                    specificURLS.Items.RemoveAt(i)
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Function removeyahoo()
        Try
            For i As Integer = specificURLS.Items.Count - 1 To 0 Step -1
                If specificURLS.Items(i).Contains("yahoo.com") Then
                    specificURLS.Items.RemoveAt(i)
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Function removecache()
        Try
            For i As Integer = specificURLS.Items.Count - 1 To 0 Step -1
                If specificURLS.Items(i).Contains("webcache.googleusercontent.com") Then
                    specificURLS.Items.RemoveAt(i)
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Function removeuploads()
        Try
            For i As Integer = specificURLS.Items.Count - 1 To 0 Step -1
                If specificURLS.Items(i).Contains("weebly.com/uploads/") Then
                    specificURLS.Items.RemoveAt(i)
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Function removegoogle()
        Try
            For i As Integer = specificURLS.Items.Count - 1 To 0 Step -1
                If specificURLS.Items(i).Contains("google.com") Then
                    specificURLS.Items.RemoveAt(i)
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Function removemicrosoft()
        Try
            For i As Integer = specificURLS.Items.Count - 1 To 0 Step -1
                If specificURLS.Items(i).Contains("go.microsoft.com") Then
                    specificURLS.Items.RemoveAt(i)
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Function removeduckduckgo()
        Try
            For i As Integer = specificURLS.Items.Count - 1 To 0 Step -1
                If specificURLS.Items(i).Contains("duckduckgo.com") Then
                    specificURLS.Items.RemoveAt(i)
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Function cleaner()
        Try
            removeduplicates()
            For i As Integer = specificURLS.Items.Count - 1 To 0 Step -1
                If Not specificURLS.Items(i).Contains("http") Then
                    specificURLS.Items.RemoveAt(i)
                End If
            Next
            removecache()
            removeduckduckgo()
            removeuploads()
            removemicrosoft()
            removebing()
            removeyahoo()
            removegoogle()
            checkstartspecific()
        Catch ex As Exception

        End Try
    End Function

    Function removeduplicates()
        Try
            Dim itemcount As Integer = specificURLS.Items.Count

            If itemcount > 1 Then
                Dim lastitem As String = specificURLS.Items(itemcount - 1)

                For index = itemcount - 2 To 0 Step -1
                    If specificURLS.Items(index) = lastitem Then
                        specificURLS.Items.RemoveAt(index)
                    Else
                        lastitem = specificURLS.Items(index)
                    End If
                Next
            End If
        Catch ex As Exception

        End Try
    End Function

    Function checkstartspecific()
        Try
            specificURLS.SelectedItem = specificURLS.Items.Item(0)
            checksrc.Navigate(specificURLS.SelectedItem.ToString())
        Catch ex As Exception

        End Try
    End Function

    Function geturls()
        Try
            For Each ClientControl As HtmlElement In specificBROWSE.Document.Links
                specificURLS.Items.Add(ClientControl.GetAttribute("href"))
                cleaner()
            Next
        Catch ex As Exception

        End Try
    End Function

    Function cleanerbing()
        Try

        Catch ex As Exception

        End Try
    End Function

    Function geturlsbing()
        Try
            For Each ClientControl As HtmlElement In bingbrowse.Document.Links
                specificURLS.Items.Add(ClientControl.GetAttribute("href"))
                cleaner()
            Next
        Catch ex As Exception

        End Try
    End Function

    Private Sub specificBROWSE_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles specificBROWSE.DocumentCompleted
        Try
            geturls()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub checksrc_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles checksrc.DocumentCompleted
        Try
            Dim dumpsrc As String = checksrc.DocumentText
            Dim currenturl As String = checksrc.Url.ToString()
            dumpsrc = checksrc.DocumentText.ToString()
            If InStr(dumpsrc, phraseTXT.Text) Then
                specificFOUND.Items.Add(currenturl)
                selectitem()
                specificURLS.Items.Remove(specificURLS.Items(0))
                checkstartspecific()
            Else
                selectitem()
                specificURLS.Items.Remove(specificURLS.Items(0))
                checkstartspecific()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub GoToToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GoToToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start(specificFOUND.SelectedItem.ToString())
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        Try
            My.Computer.Clipboard.SetText(specificFOUND.SelectedItem.ToString())
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RemoveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoveToolStripMenuItem.Click
        Try
            specificFOUND.Items.Remove(specificFOUND.SelectedItem)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            checksrc.Stop()
            specificBROWSE.Stop()
            bingbrowse.Stop()
            duckduckgobrowse.Stop()
            For i As Integer = 0 To specificURLS.Items.Count - 1
                specificURLS.SetSelected(i, False)
                specificFOUND.SetSelected(i, False)
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub bingbrowse_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles bingbrowse.DocumentCompleted
        Try
            geturlsbing()
        Catch ex As Exception

        End Try
    End Sub

    Function geturlsduck()
        Try
            For Each ClientControl As HtmlElement In duckduckgobrowse.Document.Links
                specificURLS.Items.Add(ClientControl.GetAttribute("href"))
                cleaner()
            Next
        Catch ex As Exception

        End Try
    End Function

    Private Sub duckduckgobrowse_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles duckduckgobrowse.DocumentCompleted
        Try
            geturlsduck()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Try
            If CheckBox1.Checked = True Then
                Me.TopMost = True
            Else
                Me.TopMost = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim itemcount As Integer = specificFOUND.Items.Count

            If itemcount > 1 Then
                Dim lastitem As String = specificFOUND.Items(itemcount - 1)

                For index = itemcount - 2 To 0 Step -1
                    If specificFOUND.Items(index) = lastitem Then
                        specificFOUND.Items.RemoveAt(index)
                    Else
                        lastitem = specificFOUND.Items(index)
                    End If
                Next
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Function getDocText(ByVal filepath As String) As String
        If File.Exists(filepath) AndAlso Path.GetExtension(filepath).ToUpper.Equals(".DOCX") Then
            Dim app As Application = New Application
            Dim doc As Document = app.Documents.Open(filepath)
            Dim doctxt As String = doc.Content.Text
            app.Quit()
            exportedTEXT.Text = doctxt
            Return doctxt
        Else
            Return Nothing
        End If
    End Function

    Function getdocxtxt()
        Try
            Dim docfilepath As String = wrdTXT.Text
            Dim doctext As String = getDocText(docfilepath)
        Catch ex As Exception

        End Try
    End Function

    Function splitter()
        Try
            Dim documentation As String = exportedTEXT.Text
            Dim paraspace As String = ""
            Dim splitted As String() = documentation.Split(New Char() {"."c})
            Dim newstring As String
            For Each newstring In splitted
                foundPARA.Items.Add(newstring)
            Next
        Catch ex As Exception

        End Try
    End Function

    Function removedupedoc()
        Try
            Dim itemcount As Integer = foundURLS.Items.Count

            If itemcount > 1 Then
                Dim lastitem As String = foundURLS.Items(itemcount - 1)

                For index = itemcount - 2 To 0 Step -1
                    If foundURLS.Items(index) = lastitem Then
                        foundURLS.Items.RemoveAt(index)
                    Else
                        lastitem = foundURLS.Items(index)
                    End If
                Next
            End If
        Catch ex As Exception

        End Try
    End Function

    Function removecachedoc()
        Try
            For i As Integer = foundURLS.Items.Count - 1 To 0 Step -1
                If foundURLS.Items(i).Contains("webcache.googleusercontent.com") Then
                    foundURLS.Items.RemoveAt(i)
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Function removeduckduckgodoc()
        Try
            For i As Integer = foundURLS.Items.Count - 1 To 0 Step -1
                If foundURLS.Items(i).Contains("duckduckgo.com") Then
                    foundURLS.Items.RemoveAt(i)
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Function removeuploadsdoc()
        Try
            For i As Integer = foundURLS.Items.Count - 1 To 0 Step -1
                If foundURLS.Items(i).Contains("weebly.com/uploads/") Then
                    foundURLS.Items.RemoveAt(i)
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Function removemicrosoftdoc()
        Try
            For i As Integer = foundURLS.Items.Count - 1 To 0 Step -1
                If foundURLS.Items(i).Contains("go.microsoft.com") Then
                    foundURLS.Items.RemoveAt(i)
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Function removebingdoc()
        Try
            For i As Integer = foundURLS.Items.Count - 1 To 0 Step -1
                If foundURLS.Items(i).Contains("bing.com") Then
                    foundURLS.Items.RemoveAt(i)
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Function removeyahoodoc()
        Try
            For i As Integer = foundURLS.Items.Count - 1 To 0 Step -1
                If foundURLS.Items(i).Contains("yahoo.com") Then
                    foundURLS.Items.RemoveAt(i)
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Function removegoogledoc()
        Try
            For i As Integer = foundURLS.Items.Count - 1 To 0 Step -1
                If foundURLS.Items(i).Contains("google.com") Then
                    foundURLS.Items.RemoveAt(i)
                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Function cleanerdocument()
        Try
            removedupedoc()
            For i As Integer = foundURLS.Items.Count - 1 To 0 Step -1
                If Not foundURLS.Items(i).Contains("http") Then
                    foundURLS.Items.RemoveAt(i)
                End If
            Next
            removecachedoc()
            removeduckduckgodoc()
            removeuploadsdoc()
            removemicrosoftdoc()
            removebingdoc()
            removeyahoodoc()
            removegoogledoc()
        Catch ex As Exception

        End Try
    End Function

    Function geturlsgoogledoc()
        Try
            For Each ClientControl As HtmlElement In googlebrowsedoc.Document.Links
                foundURLS.Items.Add(ClientControl.GetAttribute("href"))
                cleanerdocument()
            Next
        Catch ex As Exception

        End Try
    End Function

    Function documentscan()
        Try
            foundPARA.SelectedItem = foundPARA.Items.Item(0)
            googlebrowsedoc.Navigate("https://www.google.com/search?q=" + quoter.Text + foundPARA.SelectedItem.ToString() + quoter.Text)
            bingbrowsedoc.Navigate("http://bing.com/search?q=" + quoter.Text + foundPARA.SelectedItem.ToString() + quoter.Text)
            duckduckgobrowsedoc.Navigate("https://duckduckgo.com/?q=" + quoter.Text + foundPARA.SelectedItem.ToString() + quoter.Text)
            foundPARA.SelectedItem = foundPARA.Items.Item(0)
            foundPARA.Items.Remove(foundPARA.SelectedItem)

            If foundPARA.Items.Count > 0 Then
                checksrcdoc()
            End If
        Catch ex As Exception

        End Try
    End Function

    Function checksrcdoc()
        Try
            foundURLS.SelectedItem = foundURLS.Items.Item(0)
            srcCheckerdoc.Navigate(foundURLS.SelectedItem.ToString())
        Catch ex As Exception

        End Try
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            getdocxtxt()
            splitter()
            documentscan()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub googlebrowsedoc_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles googlebrowsedoc.DocumentCompleted
        Try
            geturlsgoogledoc()
        Catch ex As Exception

        End Try
    End Sub

    Function geturlsbingdoc()
        Try
            For Each ClientControl As HtmlElement In bingbrowsedoc.Document.Links
                foundURLS.Items.Add(ClientControl.GetAttribute("href"))
                cleanerdocument()
            Next
        Catch ex As Exception

        End Try
    End Function

    Private Sub bingbrowsedoc_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles bingbrowsedoc.DocumentCompleted
        Try
            geturlsbingdoc()
        Catch ex As Exception

        End Try
    End Sub

    Function geturlsduckduckgodoc()
        Try
            For Each ClientControl As HtmlElement In duckduckgobrowsedoc.Document.Links
                foundURLS.Items.Add(ClientControl.GetAttribute("href"))
                cleanerdocument()
            Next
        Catch ex As Exception

        End Try
    End Function

    Private Sub duckduckgobrowsedoc_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles duckduckgobrowsedoc.DocumentCompleted
        Try
            geturlsduckduckgodoc()
            documentscan()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub srcCheckerdoc_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles srcCheckerdoc.DocumentCompleted
        Try
            Dim dumpsrcdoc As String = checksrcdoc.DocumentText
            Dim currenturldoc As String = checksrcdoc.Url.ToString()
            dumpsrcdoc = checksrcdoc.DocumentText.ToString()
            If InStr(dumpsrcdoc, exportedTEXT.Text) Then
                foundDUPE.Items.Add(currenturldoc)
                foundURLS.SelectedItem = foundURLS.Items.Item(0)
                foundURLS.Items.Remove(foundURLS.SelectedItem)
                documentscan()
            Else
                selectitem()
                foundURLS.SelectedItem = foundURLS.Items.Item(0)
                documentscan()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            MsgBox("E.G - C:\location\location\file.docx", MsgBoxStyle.Information)
        Catch ex As Exception

        End Try
    End Sub
End Class