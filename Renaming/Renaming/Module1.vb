Imports System.IO
Imports System.Net.Mime.MediaTypeNames

Module Module1
    Dim SKU As New List(Of String)
    Dim Barcode As New List(Of String)
    Dim SkuFilePath As String = "‪C:\Users\PC035\Desktop\SKU.txt"
    Dim BarcodePath As String = "C:\Users\PC035\Desktop\Barcode.txt"
    Dim FolderPath As String = "C:\Users\PC035\Desktop\ImagePath"
    Sub Main()
        Fill_Lists()
        RenameFiles()
    End Sub
    Private Sub Fill_Lists()
        Console.WriteLine("Gathering SKU's")
        If File.Exists(SkuFilePath.Substring(1)) Then
            For Each line In File.ReadLines(SkuFilePath.Substring(1))
                SKU.Add(line.Substring(4))
            Next
        End If

        Console.WriteLine("Gathering Barcodes")
        If File.Exists(BarcodePath) Then
            For Each line In File.ReadLines(BarcodePath)
                Barcode.Add(line)
            Next
        End If

        Console.WriteLine("Inialized Lists")
    End Sub
    Private Sub RenameFiles()
        Dim Filenames() As String = Directory.GetFiles(FolderPath)
        Dim FilenamesList As New List(Of String)
        Dim Index As Integer = 0
        Dim Repeated As Boolean = False
        Dim RepeatedCount As Integer = 1
        Dim PreviousSKU As String = ""
        Dim SearchFileName As String = ""
        Dim RenamedFiles As Integer = 0
        Dim Log As String = ""

        FilenamesList.AddRange(Filenames)

        For i As Integer = 0 To FilenamesList.Count - 1
            FilenamesList.Item(i) = FilenamesList.Item(i).Substring(FolderPath.Count + 1)
            FilenamesList.Item(i) = FilenamesList.Item(i).Substring(0, FilenamesList.Item(i).LastIndexOf("."))
        Next

        Console.WriteLine("Renaming Files...")

        For j As Integer = 0 To FilenamesList.Count - 1
            If j > 0 Then
                If FilenamesList.Item(j).Contains(PreviousSKU) Then
                    Repeated = True
                    RepeatedCount += 1
                Else
                    Repeated = False
                    RepeatedCount = 1
                End If
            End If
            If FilenamesList.Item(j).Contains("_") Then
                SearchFileName = FilenamesList.Item(j).Substring(0, FilenamesList.Item(j).LastIndexOf("_"))
                PreviousSKU = SearchFileName
            Else
                SearchFileName = FilenamesList.Item(j)
                PreviousSKU = SearchFileName
            End If
            If SKU.Contains(SearchFileName) Then
                Index = SKU.FindIndex(Function(x As String) x.Contains(SearchFileName))
                'If SKU.Item(Index) = SearchFileName.Substring(0, SearchFileName.Count - 4) Then
                Console.WriteLine(SearchFileName.Substring(0))
                'Try
                FileSystem.Rename(FolderPath + "\" + FilenamesList.Item(j) + ".jpg", FolderPath + "\" + "a" + Barcode.Item(Index) + "_" + RepeatedCount.ToString + ".jpg")
                'PreviousSKU = Barcode.Item(Index)
                Log += vbCrLf + FolderPath + "\" + FilenamesList.Item(j) + ".jpg -> " + FolderPath + "\" + Barcode.Item(Index) + ".jpg"
                RenamedFiles += 1
                'Catch ex As Exception
                'PreviousSKU = Barcode.Item(Index)
                'End Try
            End If
            'End If
        Next
        Log += vbCrLf + "Renamed Files Count = " + RenamedFiles.ToString

        'Create a log
        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(AppDomain.CurrentDomain.BaseDirectory + "RenameLog.txt", False)
        file.WriteLine(Log)
        file.Close()
    End Sub
End Module
