Imports System.Globalization
Imports System.IO

Module Module1

    Sub Main()
        Dim mainDir = AppDomain.CurrentDomain.BaseDirectory
        'Dim mainDir = Directory.GetCurrentDirectory()
        'Dim mainDir = "C:\Terminal\"
        Dim ConnSQL, ts1, sqlstr, fl, path, buf, arr, folder, kup, kdef, d, sqlstr1, errfl, k, rs1
        ConnSQL = CreateObject("ADODB.Connection")
        ConnSQL.ConnectionString = "Provider=SQLOLEDB;Server=srv-otk;Database=otk;Trusted_Connection=yes;Integrated Security=SSPI;Persist Security Info=False"
        Dim fso, fserr, goreem
        Dim sort, sort2, nom_ctrl As String
        Dim dbupd(10000) As String
        Dim gormc = 0

        If Not Directory.Exists(mainDir & "OUT") Then
            Directory.CreateDirectory(mainDir & "OUT")
        End If

        If Not Directory.Exists(mainDir & "Arhiv") Then
            Directory.CreateDirectory(mainDir & "Arhiv")
        End If
        If Not Directory.Exists(mainDir & "err") Then
            Directory.CreateDirectory(mainDir & "err")
        End If

        ConnSQL.Open
        fso = CreateObject("Scripting.FileSystemObject")
        fserr = CreateObject("Scripting.FileSystemObject")
        errfl = fserr.OpenTextFile(mainDir & "err\err.txt", 8, True)
        fl = ""
        path = ""
        folder = fso.GetFolder(mainDir + "OUT")
        kup = 0
        kdef = 0
        d = 0
        k = 0
        For Each file In folder.Files
            If Left(file.Name, 3) = "skl" And Right(file.name, 4) = ".txt" Then
                path = file
                fl = file.name
                Exit For
            End If
        Next
        If Len(fl) < 5 Then
            MsgBox("Файл данных не найден!", vbInformation, "Ошибка")

            'System.Threading.Thread.Sleep(7000)
            Exit Sub
        End If
        Console.WriteLine("Введите номер контролера:")
        nom_ctrl = Console.ReadLine()


        ts1 = fso.OpenTextFile(path, 1, False)
        Do While Not ts1.AtEndOfStream
            buf = ts1.ReadLine
            arr = Split(buf, ";")
            sqlstr1 = "SELECT [shtr_kod], [Сорт],[sort13] FROM dbo.[Изделия] WHERE [shtr_kod]=" & arr(2)
            rs1 = ConnSQL.Execute(sqlstr1)
            If rs1.EOF = True Then

                Console.WriteLine(arr(2) & " не существует")
                errfl.WriteLine(CDate(arr(0) & " " & arr(1)) & vbTab & Now.ToShortTimeString & vbTab & arr(2) & " не существует")
                d = d + 1
                Continue Do
            End If
            sort = rs1(1).Value.ToString
            If sort = "2" Or sort = "6" Or sort = "7" Then sort = "1"
            sort2 = rs1(2).Value.ToString
            If sort2 = "" Then sort2 = sort

            If Left(arr(3), 1) = 2 Then
                sqlstr = "Update dbo.Изделия SET [DataUp] ='" & CDate(arr(0) & " " & arr(1)) & "', [NomUp] =" & Mid(arr(3), 2, 11) & ", [Sort13]=" & sort2 & ", [Control_13_nom]=" & nom_ctrl & ", [up_skl]=13 WHERE [shtr_kod]=" & arr(2)
                kup = kup + 1

            Else
                Dim nom_model As String
                Dim dt_otliv As DateTime?
                Dim nom_lit_brig As String

                nom_model = arr(5)
                nom_lit_brig = Right(arr(6), 1)

                dt_otliv = parse_arr6(arr(6))
                Dim dt_otlsql As String = If(dt_otliv.HasValue, $"'{dt_otliv.Value:yyyy-MM-dd}'", "NULL")


                If arr(4) = "" Then arr(4) = 4
                If arr(4) = 9 Then
                    gormc = gormc + 1
                    sqlstr = $"Update dbo.Изделия SET [sort13]=4, goreem=2, [NomUp]=null, [Control_13_nom]= {nom_ctrl}, [up_skl]=13, [dt_otliv]={dt_otlsql}, [НомМодели]={nom_model}, [НомЛитБригады]={nom_lit_brig}  WHERE [shtr_kod]={arr(2)}"
                    dbupd(k) = sqlstr
                    k = k + 1
                    Continue Do
                Else
                    goreem = 0
                End If
                If arr(4) = "" Then
                    sort2 = "4"
                Else
                    sort2 = arr(4)
                End If
                'sqlstr = "INSERT INTO dbo.t1 (d1,shtr,upak, def) SELECT '" & Cdate (arr(0) &" "&arr(1)) &"'," &arr(2) &", null," & arr(3)
                sqlstr = $"Update dbo.Изделия SET [DataUp] = {CDate(arr(0) & " " & arr(1))} , DefUp ={arr(3)} ,[sort13]={sort2}, goreem={goreem} ,[pereat]=1, [NomUp]=null, [Control_13_nom]={nom_ctrl}, [up_skl]=13, [dt_otliv]={dt_otliv}, [НомМодели]={nom_model}, [НомЛитБригады]={nom_lit_brig}  WHERE [shtr_kod]={arr(2)}"
                kdef = kdef + 1

            End If
            'ConnSQL.execute = sqlstr
            dbupd(k) = sqlstr
            k = k + 1
            ConnSQL.execute("Update dbo.sklad SET [13skl]='1' WHERE [shtr]=" & arr(2))

        Loop
        sqlstr = "INSERT INTO dbo.LogUpak (data,CountUp,CountBrak) SELECT getdate()," & kup & "," & kdef
        'ConnSQL.execute = sqlstr


        ConnSQL.Close


        ConnSQL.Open
        ConnSQL.BeginTrans
        For i = 0 To k - 1
            ConnSQL.Execute(dbupd(i))
            'MsgBox(dbupd(i))
        Next
        ConnSQL.CommitTrans
        ConnSQL.Close
        ts1.Close
        errfl.Close
        ts1 = fso.GetFile(path)
        Try
            ts1.move(mainDir & "Arhiv\" & fl)
        Catch ex As Exception
            ts1.delete(mainDir & "OUT\" & fl)
        End Try

        MsgBox("Упаковано: " & kup & vbNewLine & "Переатестаций: " & kdef & vbNewLine & "Не найдено: " & d & vbNewLine & "На реэмалирование:" & gormc, vbOKOnly, "Данные успешно загружены")

    End Sub

    Public Function parse_arr6(input As String) As DateTime?
        ' Проверка на null и длину
        If String.IsNullOrEmpty(input) OrElse input.Length <> 7 Then
            Return Nothing
        End If

        Dim datePart As String = input.Substring(0, 6)
        Dim dateValue As DateTime

        ' Пытаемся распарсить
        If DateTime.TryParseExact(datePart, "yyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, dateValue) Then
            Return dateValue
        End If

        Return Nothing
    End Function



End Module
