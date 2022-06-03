Imports System.Drawing
Imports System.IO
Imports System.Net.Mail
Imports System.Text

Module Module1
    Public G_AppName As String
    Public G_DirLocale As String
    Public G_LogFileName As String
    Public G_RiepFileName As String
    Public G_RiepFileName2 As String
    Public G_DirInput As String
    Public G_DirOutput As String
    Public G_Funzione As String
    Public G_List As Boolean
    Public G_StrFileName As String
    Public G_StrFileName2 As String
    Public G_StrFileNameOK As String
    Public G_StrFileNameKO As String
    Public G_FileKO As Boolean

    Sub Main()
        'InviaMail("Prova Invio da VB")
        Inizializza()

        LogFile("Main", "I", "0", "")
        LogFile("Main", "I", "0", "")
        LogFile("Main", "I", "0", "START")

        RieFile("-----------")
        RieFile("START")

        ZplConverter()


        RieFile("END")
        RieFile("----")
        RieFile("")
        LogFile("Main", "I", "0", "END")

    End Sub


    Sub Inizializza()
        Dim Opzione As String
        'Esempio di blocco che scorre un insieme (argomenti della stringa di comamndo)

        G_AppName = My.Settings.NomeApp
        G_DirLocale = ""
        G_List = False


        G_LogFileName = My.Settings.LogFile
        G_RiepFileName = My.Settings.RiepFile
        G_RiepFileName2 = G_RiepFileName + Format(Date.Now, ".yyyyMMdd_HHmm") + ".log"
        G_DirLocale = My.Settings.DirLocale
        G_DirInput = My.Settings.DirInput
        G_DirOutput = My.Settings.DirOutput
        G_Funzione = My.Settings.Funzione

        For Each argument As String In My.Application.CommandLineArgs
            Opzione = Left(argument, 2)
            Select Case Opzione
                Case "-d"
                    G_DirLocale = Mid(argument, 4)
                    G_LogFileName = G_DirLocale + G_AppName + ".log"
                    G_RiepFileName = G_DirLocale + G_AppName + "Riepilogo.log"
                    G_RiepFileName2 = G_RiepFileName + Format(Date.Now, ".yyyyMMdd_HHmm") + ".log"
                    'LogConsole("Opzione " + Opzione + " Folder -> " + G_DirLocale, "N")
                Case "-l"
                    G_List = True
                    'LogConsole("Opzione " + Opzione + " List -> " + G_List.ToString, "N")
                Case "-f"
                    G_Funzione = Mid(argument, 4)

            End Select
        Next
        'LogConsole("#", "S")

    End Sub

    Sub LogConsole(ByVal str As String, ByVal chk As String)
        Console.WriteLine(str)
        If chk = "S" Then
            Console.WriteLine("Premi un tasto per continuare")
            Console.ReadKey()
        End If

    End Sub

    Sub RieFile(ByVal str As String)
        Dim Msg As String

        Msg = Date.Now.ToString + " " + str & vbCrLf

        My.Computer.FileSystem.WriteAllText(G_RiepFileName, Msg, True)
    End Sub
    Sub RieFile2(ByVal str As String, ByVal fl As Boolean)
        Dim Msg As String

        If fl = False Then

            Msg = str & vbCrLf

            My.Computer.FileSystem.WriteAllText(G_RiepFileName2, Msg, True)
        End If

    End Sub

    Sub LogFile(ByVal strModulo As String, ByVal tipoErr As String, ByVal codeErr As String, ByVal str As String)

        Dim DescrErr As String, MsgErr As String

        Select Case tipoErr
            Case "E"
                DescrErr = " Err: "
            Case "W"
                DescrErr = " War: "
            Case "I"
                DescrErr = " Inf: "
            Case Else
                DescrErr = " YYY: "

        End Select

        MsgErr = Date.Now.ToString + DescrErr + tipoErr + codeErr + "  <" + strModulo + "> " + str & vbCrLf

        My.Computer.FileSystem.WriteAllText(G_LogFileName, MsgErr, True)

    End Sub

    Sub InviaMail(ByVal strOggetto As String)


        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential("andrea.boscaini.cc@gmail.com", "punto9000")
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.gmail.com"

            e_mail = New MailMessage()
            e_mail.From = New MailAddress("andrea.boscaini@gmail.com")
            e_mail.To.Add("andrea_boscaini@hotmail.com")
            e_mail.Subject = strOggetto
            e_mail.IsBodyHtml = False
            e_mail.Body = "Prova di invio messaggio"
            Smtp_Server.Send(e_mail)
            MsgBox("Mail Sent")

        Catch error_t As Exception
            MsgBox(error_t.ToString)
        End Try

    End Sub

    Sub ZplConverter()

        'Variabili di inizializzazione
        Dim FolderInput As String
        Dim FolderOutput As String
        Dim blackLimit As Integer
        Dim total As Integer
        Dim widthBytes As Integer
        Dim mapCode As New Dictionary(Of Integer, String)
        mapCode.Add(1, "G")
        mapCode.Add(2, "H")
        mapCode.Add(3, "I")
        mapCode.Add(4, "J")
        mapCode.Add(5, "K")
        mapCode.Add(6, "L")
        mapCode.Add(7, "M")
        mapCode.Add(8, "N")
        mapCode.Add(9, "O")
        mapCode.Add(10, "P")
        mapCode.Add(11, "Q")
        mapCode.Add(12, "R")
        mapCode.Add(13, "S")
        mapCode.Add(14, "T")
        mapCode.Add(15, "U")
        mapCode.Add(16, "V")
        mapCode.Add(17, "W")
        mapCode.Add(18, "X")
        mapCode.Add(19, "Y")
        mapCode.Add(20, "g")
        mapCode.Add(40, "h")
        mapCode.Add(60, "i")
        mapCode.Add(80, "j")
        mapCode.Add(100, "k")
        mapCode.Add(120, "l")
        mapCode.Add(140, "m")
        mapCode.Add(160, "n")
        mapCode.Add(180, "o")
        mapCode.Add(200, "p")
        mapCode.Add(220, "q")
        mapCode.Add(240, "r")
        mapCode.Add(260, "s")
        mapCode.Add(280, "t")
        mapCode.Add(300, "u")
        mapCode.Add(320, "v")
        mapCode.Add(340, "w")
        mapCode.Add(360, "x")
        mapCode.Add(380, "y")
        mapCode.Add(400, "z")

        'Regolazione limiti del nero
        blackLimit = 50 * 768 / 100

        'Percorso cartelle
        FolderInput = G_DirInput
        FolderOutput = G_DirOutput

        'Lettura di tutte le immagini nella directory
        Dim directory As New IO.DirectoryInfo(FolderInput)
        If directory.Exists Then
            Dim files() As IO.FileInfo = directory.GetFiles("*.*")
            For Each file As IO.FileInfo In files
                If file.Exists Then
                    Dim image As New Bitmap(file.FullName)

                    'Funzione per conversione immagine in binary e Hex
                    Using image
                        Dim stringBuilder As New StringBuilder()
                        Dim height As Integer = image.Height
                        Dim width As Integer = image.Width
                        Dim red, green, blue, index As Integer
                        Dim auxBinaryChar() As Char = {"0", "0", "0", "0", "0", "0", "0", "0"}

                        'Calcolo lunghezza della riga in bytes
                        If (width Mod 8) > 0 Then
                            widthBytes = ((width / 8) + 1)
                        Else
                            widthBytes = width / 8
                        End If
                        total = widthBytes * height

                        'Calcolo valore pixel in binario
                        For h As Integer = 0 To height - 1
                            For w As Integer = 0 To width - 1
                                Dim pixelColor As Color = image.GetPixel(w, h)
                                red = pixelColor.R
                                green = pixelColor.G
                                blue = pixelColor.B
                                Dim auxChar As Char = "1"
                                Dim totalColor As Integer = red + green + blue
                                If totalColor > blackLimit Then
                                    auxChar = "0"
                                End If
                                auxBinaryChar(index) = auxChar
                                index += 1
                                Dim fourByteBinary As String

                                'Conversione binary -> Hex
                                If index = 8 Or w = (width - 1) Then
                                    Dim decimals As Integer = Convert.ToInt32(auxBinaryChar, 2)
                                    If (decimals > 15) Then
                                        fourByteBinary = Hex(decimals)
                                    Else
                                        fourByteBinary = "0" + Hex(decimals)
                                    End If
                                    stringBuilder.Append(fourByteBinary)
                                    auxBinaryChar = {"0", "0", "0", "0", "0", "0", "0", "0"}
                                    index = 0
                                End If

                            Next
                            'Fine riga
                            stringBuilder.AppendLine()
                        Next

                        'Funzione di conversione da Hex a Zpl
                        Dim code As String = stringBuilder.ToString.Replace(vbCr, " ").Replace(vbLf, "")
                        Dim maxlinea As Integer = widthBytes * 2
                        Dim sbCode As New StringBuilder()
                        Dim sbLinea As New StringBuilder()
                        Dim previousLine As String = Nothing
                        Dim counter As Integer = 1
                        Dim aux As Char = code.First
                        Dim firstChar As Boolean = False

                        'Lettura codice Hex
                        For i As Integer = 1 To code.Length - 1
                            If firstChar Then
                                aux = code.Chars(i)
                                firstChar = False
                                Continue For
                            End If

                            'Gestione cambio riga
                            If code.Chars(i).Equals(" "c) Then
                                If counter >= maxlinea And aux = "0" Then
                                    sbLinea.Append(",")
                                ElseIf counter >= maxlinea And aux = "F" Then
                                    sbLinea.Append("!")
                                ElseIf counter > 20 Then
                                    Dim multi20 As Integer = (counter \ 20) * 20
                                    Dim resto20 As Integer = (counter Mod 20)
                                    sbLinea.Append(mapCode(multi20))
                                    If resto20 <> 0 Then
                                        sbLinea.Append(mapCode(resto20)).Append(aux)
                                    Else
                                        sbLinea.Append(aux)
                                    End If
                                Else
                                    sbLinea.Append(mapCode(counter)).Append(aux)
                                End If
                                counter = 1
                                firstChar = True

                                'Gestione riga uguale a quella di prima
                                If sbLinea.ToString().Equals(previousLine) Then
                                    sbCode.Append(":")
                                Else
                                    sbCode.Append(sbLinea)
                                End If
                                previousLine = sbLinea.ToString()
                                sbLinea.Length = 0
                                Continue For
                            End If

                            'Gestione carattere = 0 e <> 0
                            If aux = code.Chars(i) Then
                                counter += 1
                            Else
                                'Else if da risolvere
                                If counter > 20 Then
                                    Dim multi20 As Integer = (counter \ 20) * 20
                                    Dim resto20 As Integer = (counter Mod 20)
                                    sbLinea.Append(mapCode(multi20))
                                    If resto20 <> 0 Then
                                        sbLinea.Append(mapCode(resto20)).Append(aux)
                                    Else
                                        sbLinea.Append(aux)
                                    End If
                                Else
                                    sbLinea.Append(mapCode(counter)).Append(aux)
                                End If
                                counter = 1
                                aux = code.Chars(i)
                            End If
                        Next

                        'Riga di spazio
                        sbCode.AppendLine()

                        'Nome del file di testo preso dal nome dell'immagine
                        Dim filename As String = Path.GetFileNameWithoutExtension(file.FullName) + ".txt"

                        'Riga di output elaborata secondo ZPL
                        Dim output As String = total.ToString + "," + total.ToString + "," + widthBytes.ToString + "," + sbCode.ToString

                        'Stampa su console del codice per immagine etichetta
                        'Console.WriteLine(file.FullName)
                        'Console.WriteLine(total.ToString + "," + total.ToString + "," + widthBytes.ToString + "," + sbCode.ToString)

                        'Stampa su file del codice per immagine etichetta
                        My.Computer.FileSystem.WriteAllText(G_DirOutput + filename, output, True)

                    End Using
                End If
            Next
        End If
    End Sub
End Module
