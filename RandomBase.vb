Imports System.IO

Module RandomBase
    Structure Record
        <VBFixedString(15)> Public FirstName As String
        <VBFixedString(15)> Public LastName As String
        <VBFixedString(35)> Public Company As String
        <VBFixedString(35)> Public Address As String
        <VBFixedString(30)> Public City As String
        <VBFixedString(25)> Public County As String
        <VBFixedString(2)> Public State As String
        <VBFixedString(5)> Public Zip As String
        <VBFixedString(12)> Public Phone As String
        <VBFixedString(12)> Public Cell As String
        <VBFixedString(30)> Public Email As String
        <VBFixedString(40)> Public Web As String
    End Structure


    Sub Main()
        Dim menuChoice As ConsoleKeyInfo
        Call OpenMenu()
        menuChoice = Console.ReadKey()
        Select Case menuChoice.Key
            Case ConsoleKey.F1
                Call CreateFile()
            Case ConsoleKey.F2
                Call RetrieveRecords()
            Case ConsoleKey.F3
                Call AddRecord()
                ' Console.WriteLine(menuChoice.Key.ToString)
            Case ConsoleKey.F4
                Console.WriteLine(menuChoice.Key.ToString)
            Case ConsoleKey.F5
                Console.WriteLine(menuChoice.Key.ToString)
            Case ConsoleKey.F6
                Console.WriteLine(menuChoice.Key.ToString)
            Case Else
                Console.WriteLine("not a function key")
        End Select
        Console.Write("press any key to continue...")
        Console.ReadKey(False)
    End Sub

    Sub OpenMenu()
        Console.Clear()
        Console.BackgroundColor = ConsoleColor.DarkBlue
        Console.ForegroundColor = ConsoleColor.White
        Call StarLine()
        Call MenuLine(" random access file manager main menu")
        Call MenuLine(" F1 = create")
        Call MenuLine(" F2 = retrieve")
        Call MenuLine(" F3 = add")
        Call MenuLine(" F4 = modify")
        Call MenuLine(" F5 = delete")
        Call MenuLine(" F6 = terminate")
        Call StarLine()
        Console.ResetColor()
    End Sub

    Sub FieldMenu()
        Console.Clear()
        Console.BackgroundColor = ConsoleColor.DarkBlue
        Console.ForegroundColor = ConsoleColor.White
        Call StarLine()
        Call MenuLine(" random access file record field menu")
        Call MenuLine(" 1 - first name")
        Call MenuLine(" 2 = last name")
        Call MenuLine(" 3 = company")
        Call MenuLine(" 4 = address")
        Call MenuLine(" 5 = city")
        Call MenuLine(" 6 = county")
        Call MenuLine(" 7 = state")
        Call MenuLine(" 8 = zip code")
        Call MenuLine(" 9 = telephone no.")
        Call MenuLine(" 10 = cellular no.")
        Call MenuLine(" 11 = e-mail address")
        Call MenuLine(" 12 = web page url")
        Call StarLine()
        Console.ResetColor()
    End Sub

    Sub StarLine()
        For i = 1 To 80
            Console.Write("*")
        Next
    End Sub

    Sub MenuLine(str As String)
        Console.Write("* ")
        If (str.Length >= 77) Then
            str = str.Substring(0, 76)
        End If
        Console.Write(str.PadRight(77, " "))
        Console.Write("*")
    End Sub

    Sub CreateFile()
        Dim filename As String
        Dim person As Record
        Dim card(12) As String
        Dim i, RecNo As Integer

        Console.WriteLine("enter .CSV file to convert below")
        filename = Console.ReadLine()
        Console.Write(filename)

        FileOpen(1, "random.dat", OpenMode.Random, , , 256)
        RecNo = 1
        Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(filename)
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")
            Dim currentRow As String()
            currentRow = MyReader.ReadFields()
            While Not MyReader.EndOfData
                Try
                    currentRow = MyReader.ReadFields()
                    Dim currentField As String
                    i = 1
                    For Each currentField In currentRow
                        card(i) = currentField
                        i = i + 1
                    Next
                    Call CardtoRecord(person, card)
                    FilePut(1, person, RecNo)
                    RecNo = RecNo + 1
                Catch ex As Microsoft.VisualBasic.
                            FileIO.MalformedLineException
                    Console.Write("Line " & ex.Message &
                    "is not valid and will be skipped.")
                End Try
            End While
        End Using
        FileClose(1)
    End Sub

    Sub CardtoRecord(ByRef person As Record, ByRef card() As String)
        person.FirstName = card(1)
        person.LastName = card(2)
        person.Company = card(3)
        person.Address = card(4)
        person.City = card(5)
        person.County = card(6)
        person.State = card(7)
        person.Zip = card(8)
        person.Phone = card(9)
        person.Cell = card(10)
        person.Email = card(11)
        person.Web = card(12)
    End Sub

    Sub RetrieveRecords()
        Dim person As Record
        Dim SearchFlds(12), FieldNo, RecNo, tgtidx, startidx As Integer
        Dim reply, ans, target As String

        Dim FileName As String
        Console.Write("Enter random access file : ")
        FileName = Console.ReadLine()

        FileOpen(1, FileName, OpenMode.Random, OpenAccess.Read, , 256)
        Dim info As New FileInfo(FileName)
        Dim n As Integer = info.Length / 256
        If (n <= 0) Then
            FileClose(1)
            System.IO.File.Delete(FileName)
            Console.Write(" empty file - processing terminated - press any key to continue . . .")
            Console.ReadKey(True)
            End
        End If

        ans = "y"
        Call FieldMenu()
        Dim fieldnumbers As String()
        Dim fieldnumber As String

        Console.WriteLine("search fields")
        Call GetFieldChoices(fieldnumbers, SearchFlds, FieldNo)

        Dim idx(n) As Integer
        Dim field(n) As String
        Dim fld As String
        Dim SearchKey As String
        For RecNo = 1 To n
            FileGet(1, person, RecNo)
            SearchKey = ""
            For i = 0 To FieldNo - 2
                Call GetField(fld, person, SearchFlds(i))
                SearchKey = Trim(SearchKey) & Trim(fld) & ","
            Next
            Call GetField(fld, person, SearchFlds(FieldNo - 1))
            SearchKey = Trim(SearchKey) & Trim(fld)
            field(RecNo) = SearchKey
        Next
        Call BuildIndex(idx, field, n)
        Console.Write("Enter search target field string : ")
        target = Console.ReadLine()
        tgtidx = BinarySearch(target, field, idx, n)
        If (tgtidx = 0) Then Return
        RecNo = idx(tgtidx)
        startidx = tgtidx
        FileGet(1, person, RecNo)
        Do While ((ans = "y" Or ans = "Y") And String.Compare(RTrim(target), RTrim(field(RecNo))) = 0)
            RecNo = idx(tgtidx)
            Call ShowRecord(person, RecNo)
            Console.Write("continue browsing records? (y/n) ")
            ans = Console.ReadLine()
            tgtidx = tgtidx + 1
            RecNo = idx(tgtidx)
            FileGet(1, person, RecNo)
        Loop
        Call Report(idx, field, startidx, target, n)
        FileClose(1)
    End Sub

    Sub Report(ByRef idx() As Integer, ByRef field() As String, ByVal tgtidx As Integer, ByRef target As String, ByVal n As Integer)
        Dim reply, attribute, selectrow, orderrow, fieldnumbers() As String
        Dim i, j, k, RecNo, fldidx(12) As Integer, sortidx(12) As Integer
        Dim person As Record

        ' get select fields info from user
        Call FieldMenu()
        Console.WriteLine("select fields")
        Call GetFieldChoices(fieldnumbers, fldidx, j)

        ' get order by choice from user
        Console.WriteLine("sort fields")
        Call GetFieldChoices(fieldnumbers, sortidx, k)

        Console.Clear()
        Console.BackgroundColor = ConsoleColor.DarkBlue
        Console.ForegroundColor = ConsoleColor.White

        RecNo = idx(tgtidx)
        FileGet(1, person, RecNo)

        Dim table As New List(Of String)
        ' we shall store the order by keys of the relation in sortkeys
        Dim sortkeys As New List(Of String)

        Dim m As Integer = 0
        Do While (String.Compare(RTrim(target), RTrim(field(RecNo))) = 0)
            ' grab the tuple - aka the 'select' fields in SQL
            selectrow = ""
            For i = 0 To j - 1
                Call GetField(attribute, person, fldidx(i))
                selectrow = selectrow & attribute
                If (fldidx(i) >= 7 And fldidx(i) <= 10) Then
                    selectrow = selectrow & " "
                End If
            Next i
            table.Add(selectrow)

            ' grab the mth sorting key
            orderrow = ""
            For i = 0 To k - 1
                Call GetField(attribute, person, sortidx(i))
                orderrow = orderrow & attribute
            Next i
            sortkeys.Add(orderrow)

            m = m + 1
            tgtidx = tgtidx + 1
            RecNo = idx(tgtidx)
            FileGet(1, person, RecNo)
        Loop

        Dim ordidx(m) As Integer
        Dim orderkeys(m) As String

        For i = 1 To m
            orderkeys(i) = sortkeys.Item(i - 1)
        Next

        Call BuildIndex(ordidx, orderkeys, m)

        Call StarLine()

        Dim relation As String() = table.ToArray
        Dim row As String

        For i = 1 To m
            row = ""
            'Call MenuLine(table.Item(ordidx(i) - 1))
            Call MenuLine(relation(ordidx(i) - 1))
            If ((i + 1) Mod 21 = 0) Then
                Call StarLine()
                Console.ResetColor()
                Console.Write("press any key to continue . . .")
                Console.ReadKey(False)
                Console.Clear()
                Console.BackgroundColor = ConsoleColor.DarkBlue
                Console.ForegroundColor = ConsoleColor.White
                Call StarLine()
            End If
        Next i

        Call StarLine()
        Console.ResetColor()
        Console.WriteLine(" m = {0}", m)
    End Sub

    Sub GetField(ByRef field As String, ByRef person As Record, ByVal FieldNo As Integer)
        Select Case FieldNo
            Case 1
                field = person.FirstName
            Case 2
                field = person.LastName
            Case 3
                field = person.Company
            Case 4
                field = person.Address
            Case 5
                field = person.City
            Case 6
                field = person.County
            Case 7
                field = person.State
            Case 8
                field = person.Zip
            Case 9
                field = person.Phone
            Case 10
                field = person.Cell
            Case 11
                field = person.Email
            Case 12
                field = person.Web
            Case Else
                field = person.FirstName + person.LastName
        End Select
    End Sub

    Sub ShowRecord(ByRef person As Record, ByVal RecNo As Integer)
        Console.Clear()
        Console.WriteLine("record no. " + RecNo.ToString)
        Console.BackgroundColor = ConsoleColor.DarkBlue
        Console.ForegroundColor = ConsoleColor.White
        Call StarLine()
        Call MenuLine(person.FirstName)
        Call MenuLine(person.LastName)
        Call MenuLine(person.Company)
        Call MenuLine(person.Address)
        Call MenuLine(person.City)
        Call MenuLine(person.County)
        Call MenuLine(person.State)
        Call MenuLine(person.Zip)
        Call MenuLine(person.Phone)
        Call MenuLine(person.Cell)
        Call MenuLine(person.Email)
        Call MenuLine(person.Web)
        Call StarLine()
        Console.ResetColor()
    End Sub

    Sub BuildIndex(ByRef idx() As Integer, ByRef field() As String, ByVal n As Integer)
        Dim i, j, inc, tmpidx As Integer
        Dim tmp As String

        For i = 1 To n
            idx(i) = i
        Next i
        inc = Math.Max(1, n * 5 / 11)
        Do While (inc > 1)
            inc = Math.Max(1, inc * 5 / 11)
            For i = inc + 1 To n
                tmpidx = idx(i)
                tmp = field(tmpidx)
                j = i
                Do While (String.Compare(field(idx(j - inc)), tmp) > 0)
                    idx(j) = idx(j - inc)
                    j = j - inc
                    If (j <= inc) Then Exit Do
                Loop
                idx(j) = tmpidx
            Next i
        Loop
    End Sub

    Function BinarySearch(ByRef target As String, ByRef field() As String, ByRef idx() As Integer,
                     ByVal n As Integer) As Integer
        Dim tgtidx, first, last, mid As Integer

        tgtidx = 1
        first = 1
        last = n
        mid = (first + last) \ 2
        Do While (first < last)
            If (String.Compare(target, field(idx(mid))) > 0) Then
                first = mid + 1
            Else
                last = mid
            End If
            mid = (first + last) \ 2
            If (mid >= last) Then Exit Do
        Loop
        If (first = last And String.Compare(RTrim(field(idx(mid))), RTrim(target)) = 0) Then
            BinarySearch = mid
        Else
            BinarySearch = 0
        End If
    End Function

    Sub AddRecord()
        Dim n As Integer
        Dim ans As String
        Dim person As Record

        n = 1
        FileOpen(1, "random.dat", OpenMode.Random, , , 256)
        Do While Not EOF(1)
            FileGet(1, person, n)
            n = n + 1
        Loop
        Call FieldMenu()
        Console.Write(" First Name : ")
        person.FirstName = Console.ReadLine()
        Console.Write(" Last Name : ")
        person.LastName = Console.ReadLine()
        Console.Write(" Organization : ")
        person.Company = Console.ReadLine()
        Console.Write(" Street Address : ")
        person.Address = Console.ReadLine()
        Console.Write(" City : ")
        person.City = Console.ReadLine()
        Console.Write(" County : ")
        person.County = Console.ReadLine()
        Console.Write(" State : ")
        person.State = Console.ReadLine()
        Console.Write(" Zip Code : ")
        person.Zip = Console.ReadLine()
        Console.Write(" Telephone No. : ")
        person.Phone = Console.ReadLine()
        Console.Write(" Cellular No. : ")
        person.Cell = Console.ReadLine()
        Console.Write(" E-mail Address : ")
        person.Email = Console.ReadLine()
        Console.Write(" Web Page Url : ")
        person.Web = Console.ReadLine()
        Call ShowRecord(person, n)
        Console.Write(" Write record to file? (y/n) : ")
        ans = Console.ReadLine()
        If (ans = "y" Or ans = "Y") Then
            FilePut(1, person, n)
        End If
        FileClose(1)
    End Sub

    Sub GetFieldChoices(ByRef FieldNumbers() As String, ByRef FieldIdx() As Integer, ByRef NoFlds As Integer)
        Dim i As Integer
        Dim fieldnumber, reply As String
        Console.Write("enter field choices e. g. (1 5 7) ? ")
        reply = Console.ReadLine()
        FieldNumbers = reply.Split(New Char() {" "c})
        NoFlds = 0
        For Each fieldnumber In FieldNumbers
            FieldIdx(NoFlds) = Integer.Parse(fieldnumber)
            NoFlds = NoFlds + 1
        Next
    End Sub
End Module
