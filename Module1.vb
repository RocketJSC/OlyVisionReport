Imports System.IO

Module Module1
    Public arPriestTarget() As String
    Public arTargetPriest() As String

    Public arSkillCity() As String
    Public arCitySkill() As String

    Sub Main()

        'Dim files() As String = IO.Directory.GetFiles("\\tsclient\C\Users\mark\Dropbox\Olympia\Reports")
        Dim arPlayerList() = {"ANDREW", "CHRIS", "JOHNNY", "KELLY", "MARCO", "OLEG", "OSSWID", "ROB", "TIM"}

        Dim strFileLocation As String = "C:\Users\mark\Dropbox\Olympia\Reports"
        Console.Write("Enter File Location [C:\Users\mark\Dropbox\Olympia\Reports]: ")
        Dim strFileLocEntry As String = Console.ReadLine()
        If Trim(strFileLocEntry) <> "" Then
            strFileLocation = strFileLocEntry
        End If
        Console.Write("Enter Turn: ")
        Dim strTurn As String = Console.ReadLine()

        Console.WriteLine("Retrieving list of files to process ...")

        Dim myDirectory As New IO.DirectoryInfo(strFileLocation)
        Dim files() As String
        Try
            files = myDirectory.GetFiles.OrderBy(Function(x) x.Extension).Select(Function(x) x.FullName).ToArray
        Catch ex As Exception
            Console.WriteLine("File error " + ex.Message)
            Console.WriteLine("Press <Enter> to terminate program.")
            Console.ReadLine()
            End
        End Try

        For Each filename As String In files
            If (InStr(filename, ".0") > 0 Or InStr(filename, ".1") > 0) Then
                If InStr(UCase(filename), ".O.") = 0 Then
                    For i As Integer = 0 To UBound(arPlayerList)
                        If InStr(UCase(filename), arPlayerList(i)) > 0 Then
                            Console.WriteLine(filename)
                            Call Process_File(filename)
                        End If
                    Next i
                End If
            End If
        Next
        Call Print_Vision_Report(strTurn, strFileLocation)
        'Call Print_Skill_Report(strTurn, strFileLocation)
        'Array.Sort(arSkillCity)
        'For index = 0 To UBound(arSkillCity) - 1
        '    Console.WriteLine(arSkillCity(index))
        'Next
        Console.WriteLine("Report finished ...")
        Console.ReadLine()

    End Sub
    Private Sub Print_Vision_Report(strTurn As String, strFileLocation As String)
        Dim outputfile As New StreamWriter(strFileLocation & "\visions" & "." & Trim(strTurn))

        outputfile.WriteLine("** Visions by Priest/Target **")

        Array.Sort(arPriestTarget)
        Dim savePriest As String = ""
        Dim strOutLine As String = ""
        Dim strNewLine As String = ""
        For i As Integer = 0 To UBound(arPriestTarget) - 1
            Dim strSplit() As String = Split(arPriestTarget(i), "|")
            If savePriest <> strSplit(0) Then
                outputfile.WriteLine(strOutLine)
                outputfile.WriteLine("")
                savePriest = strSplit(0)
                strOutLine = strSplit(1) & " " & strSplit(0) & ":" & " " & strSplit(3) & " " & strSplit(2) & " " & strSplit(4)
            Else
                strNewLine = strSplit(3) & " " & strSplit(2) & strSplit(4)
                If Len(strOutLine) + Len(strNewLine) > 80 Then
                    outputfile.WriteLine(strOutLine & ",")
                    strOutLine = strNewLine
                Else
                    strOutLine = strOutLine & ", " & strNewLine
                End If
            End If
        Next
        outputfile.WriteLine(strOutLine)

        outputfile.WriteLine("")
        outputfile.WriteLine("** Visions by Target/Priest **")

        Array.Sort(arTargetPriest)
        Dim saveTarget As String = ""
        strOutLine = ""
        For i As Integer = 0 To UBound(arPriestTarget) - 1
            Dim strSplit() As String = Split(arTargetPriest(i), "|")
            If savePriest <> strSplit(0) Then
                outputfile.WriteLine(strOutLine)
                outputfile.WriteLine("")
                savePriest = strSplit(0)
                strOutLine = strSplit(1) & " " & strSplit(0) & ":" & " " & strSplit(3) & " " & strSplit(2) & " " & strSplit(4)
            Else
                strNewLine = strSplit(3) & " " & strSplit(2) & strSplit(4)
                If Len(strOutLine) + Len(strNewLine) > 80 Then
                    outputfile.WriteLine(strOutLine & ",")
                    strOutLine = strNewLine
                Else
                    strOutLine = strOutLine & ", " & strNewLine
                End If
            End If
        Next
        outputfile.WriteLine(strOutLine)

        outputfile.Close()
    End Sub
    Private Sub Process_File(filename As String)
        Dim sr As StreamReader = New StreamReader(filename)
        Dim InLine As String
        Dim intA As Integer
        Dim strPriest As String
        Dim strVisionTarget As String
        Dim strCityId As String
        Dim strProvinceId As String
        Do While sr.Peek() >= 0
            InLine = sr.ReadLine()
            ' Vision report area
            If InStr(InLine, "] receives a vision of ") > 0 Then
                intA = InStr(InLine, "] receives a vision of ")
                If InStr(Len(Trim(InLine)) - 2, InLine, "]:") = 0 _
                    And InStr(Len(Trim(InLine)) - 2, InLine, "].") = 0 Then
                    Dim Inline2 As String = sr.ReadLine()
                    InLine = InLine & Mid(Trim(Inline2), 4)
                    intA = InStr(InLine, "] receives a vision of ")
                End If
                strPriest = Trim(Mid(InLine, 5, intA - 4))
                If InStr(Len(Trim(InLine)) - 2, InLine, "]:") > 0 _
                    And InStr(Len(Trim(InLine)) - 2, InLine, "].") = 0 Then
                    strVisionTarget = Trim(Mid(InLine, intA + 23, Len(InLine) - intA - 23))
                    If arPriestTarget Is Nothing Then
                        ReDim arPriestTarget(0)
                        ReDim arTargetPriest(0)
                    Else
                        ReDim Preserve arPriestTarget(UBound(arPriestTarget) + 1)
                        ReDim Preserve arTargetPriest(UBound(arTargetPriest) + 1)
                    End If
                    arPriestTarget(UBound(arPriestTarget)) =
                        Mid(strPriest, InStr(strPriest, "["), Len(strPriest) - InStr(strPriest, "[") + 1) &
                        "|" &
                        Mid(strPriest, 1, InStr(strPriest, "[") - 2) &
                        "|" &
                        Mid(strVisionTarget, InStr(strVisionTarget, "["), Len(strVisionTarget) - InStr(strVisionTarget, "[") + 1) &
                        "|" &
                        Mid(strVisionTarget, 1, InStr(strVisionTarget, "[") - 2) &
                        "|" &
                        "(T" & Right(filename, 3) & ")"
                    arTargetPriest(UBound(arTargetPriest)) =
                        Mid(strVisionTarget, InStr(strVisionTarget, "["), Len(strVisionTarget) - InStr(strVisionTarget, "[") + 1) &
                        "|" &
                        Mid(strVisionTarget, 1, InStr(strVisionTarget, "[") - 2) &
                        "|" &
                        Mid(strPriest, InStr(strPriest, "["), Len(strPriest) - InStr(strPriest, "[") + 1) &
                        "|" &
                        Mid(strPriest, 1, InStr(strPriest, "[") - 2) &
                        "|" &
                        "(T" & Right(filename, 3) & ")"
                End If
            Else
                If InStr(InLine, " city,") > 0 Then
                    Dim strSplit() = Split(InLine, ",")
                    If strSplit.Length >= 3 Then
                        If InStr(strSplit(2), "in province") > 0 Then
                            If InStr(strSplit(0), ":") > 0 Then
                                strCityId = Right(strSplit(0), Len(strSplit(0)) - InStr(strSplit(0), ":")).Trim
                            Else
                                strCityId = strSplit(0).Trim
                            End If
                            strProvinceId = Right(strSplit(2),6)
                            InLine = sr.ReadLine()
                            Do While sr.EndOfStream = False
                                If InStr(InLine, "Skills taught here:") > 0 Then
                                    InLine = sr.ReadLine()
                                    Do While Len(InLine.Trim) > 0
                                        'If InStr(InLine, ":") > 0 Then
                                        '    InLine = Right(InLine, len(InLine) - InStr(InLine, ":"))
                                        'End If
                                        Dim strSkills = Split(InLine, ",")
                                        For index = 0 To strSkills.Length - 1
                                            If strSkills(index).Trim <> "" Then
                                                If arSkillCity Is Nothing Then
                                                    ReDim arSkillCity(0)
                                                    ReDim arCitySkill(0)
                                                    arSkillCity(UBound(arSkillCity)) = strSkills(index).Trim + "|" + strProvinceId + "|" + strCityId
                                                    arCitySkill(UBound(arCitySkill)) = strProvinceId + "|" + strCityId + "|" + strSkills(index).Trim
                                                Else
                                                    Dim index1 As Integer
                                                    For index1 = 0 To arSkillCity.Length - 1
                                                        If arCitySkill(index1) = strProvinceId + "|" + strCityId + "|" + strSkills(index).Trim Then
                                                            Exit Do
                                                        End If
                                                    Next
                                                    If index1 = arSkillCity.Length Then
                                                        ReDim Preserve arSkillCity(UBound(arSkillCity) + 1)
                                                        ReDim Preserve arCitySkill(UBound(arCitySkill) + 1)
                                                        arSkillCity(UBound(arSkillCity)) = strSkills(index).Trim + "|" + strProvinceId + "|" + strCityId
                                                        arCitySkill(UBound(arCitySkill)) = strProvinceId + "|" + strCityId + "|" + strSkills(index).Trim
                                                    End If
                                                End If
                                            End If
                                        Next
                                        InLine = sr.ReadLine()
                                    Loop
                                Else
                                    InLine = sr.ReadLine()
                                End If
                            Loop
                            'Skills taught here:
                        End If
                    End If
                Else
                        ' scrolls/potions
                    End If
            End If
        Loop
    End Sub

End Module
