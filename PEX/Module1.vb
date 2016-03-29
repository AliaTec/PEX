
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO


Module Module1

    Private Sub ReadAllText(ByVal con As SqlConnection)
        Dim cadena As String
        Dim TipoC As String
        Dim Fecha As String = "20150309" 'String.Format("{0:yyyyMMdd}", DateTime.Now)
        Dim Path As String = "C:\Users\Aliatec-HP\Documents\Victor\PEX\Ejemplo\"
        Dim GCC As String = "Z05"
        Dim LCC As String = "MX012"
        Dim PayGroup As String = "JJ"
        Dim FileName As String
        Dim Number As String = ""
        con.Open()
        Dim cmdQuery As New SqlCommand("select CASE WHEN LEN(CAST((REPLACE(SUBSTRING(Name,LEN(Name)-4,LEN(Name)),'.csv','')+1) as VARCHAR))<2 THEN '00' ELSE '0' End + CAST((REPLACE(SUBSTRING(Name,LEN(Name)-4,LEN(Name)),'.csv','')+1) as VARCHAR) as Number from PexFiles WHERE CONVERT(VARCHAR,Fecha,103)=CONVERT(VARCHAR,GETDATE(),103) and CONVERT(VARCHAR,Fecha,103)=(select MAX(CONVERT(VARCHAR,Fecha,103)) from PEXFiles)", con)
        Dim readerObj As SqlClient.SqlDataReader = cmdQuery.ExecuteReader

        While readerObj.Read
            Number = readerObj("Number").ToString
        End While
        readerObj.Close()

        If Number = "" Then
            Number = "001"
        End If

        FileName = "WAITINGEVENTS_" + GCC + "_" + LCC + "_" + PayGroup + "_D" + Fecha + "_O" + Number + ".csv"


        Dim cmd As New SqlCommand("spi_TaskRecord_NewBlock_PEX", con)
        cmd.CommandType = CommandType.StoredProcedure

        Dim Data As SqlParameter = cmd.Parameters.Add("Data", SqlDbType.VarChar, 8000)
        Data.Direction = ParameterDirection.Input
        Dim Tipo As SqlParameter = cmd.Parameters.Add("Tipo", SqlDbType.VarChar, 10)
        Tipo.Direction = ParameterDirection.Input
        Dim Archivo As SqlParameter = cmd.Parameters.Add("Archivo", SqlDbType.VarChar, 500)
        Archivo.Direction = ParameterDirection.Input

        Dim ary(0) As String


        If System.IO.File.Exists(Path + FileName) = True Then
            Dim sr As New System.IO.StreamReader(Path + FileName)

            ' Hold the amount of lines already read in a 'counter-variable' 
            Dim contador As Integer = 0
            Do While sr.Peek <> -1 ' Is -1 when no data exists on the next line of the CSV file
                ReDim Preserve ary(contador)
                ary(contador) = sr.ReadLine
                cadena = ary(contador)

                If contador = 0 Then
                    TipoC = "H"
                Else
                    TipoC = "D"
                End If

                If sr.EndOfStream Then
                    TipoC = "F"
                End If
                contador += 1

                Try
                    Data.Value = cadena
                    Tipo.Value = TipoC
                    Archivo.Value = FileName
                    cmd.ExecuteNonQuery()
                    'MsgBox("Succesfully Added", MsgBoxStyle.Information, "add")

                Catch ex As Exception
                    MsgBox(ex.Message)
                    con.Close()
                End Try

            Loop
        End If
        con.Close()

    End Sub

    Sub Main()
        Dim con As New SqlConnection("Data Source=ALIATEC; Initial Catalog=V5Pruebas; Integrated Security=True")
        ReadAllText(con)
    End Sub

End Module
