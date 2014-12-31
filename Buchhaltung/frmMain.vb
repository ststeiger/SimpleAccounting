
Public Class frmMain

    Protected m_bDontSaveWhileLoading = True
    Protected m_FileName As String = GetFileName()
    Protected m_dt As System.Data.DataTable = GetTable()


    Public Function GetTable() As System.Data.DataTable
        Dim dt As System.Data.DataTable = New System.Data.DataTable
        dt.TableName = "Contabilità"

        dt.Columns.Add("DATA", GetType(DateTime)) ' Datum
        dt.Columns.Add("DETTAGLIATO", GetType(String)) ' Details
        dt.Columns.Add("ENTRATA", GetType(Double)) ' Einnahmen
        dt.Columns.Add("USCITA", GetType(Double)) ' Ausgaben

        ' http://msdn.microsoft.com/en-us/library/system.data.datatable_events%28v=vs.110%29.aspx
        AddHandler dt.ColumnChanged, AddressOf Table_ColumnChanged
        AddHandler dt.RowChanged, AddressOf Table_RowChanged
        AddHandler dt.RowDeleted, AddressOf Table_RowDeleted
        AddHandler dt.TableCleared, AddressOf Table_Cleared
        AddHandler dt.TableNewRow, AddressOf Table_NewRow

        Return dt
    End Function


    Private Sub Table_ColumnChanged(sender As Object, e As DataColumnChangeEventArgs)
        Console.WriteLine("Column_Changed Event")

        RecalculateTotal()
    End Sub


    Private Sub Table_RowChanged(sender As Object, e As DataRowChangeEventArgs)
        Console.WriteLine("Row Changed Event")

        RecalculateTotal()
    End Sub


    Private Sub Table_RowDeleted(sender As Object, e As DataRowChangeEventArgs)
        Console.WriteLine("Row_Deleted Event")

        RecalculateTotal()
    End Sub


    Private Sub Table_Cleared(sender As Object, e As DataTableClearEventArgs)
        Console.WriteLine("Table_Cleared Event")

        RecalculateTotal()
    End Sub ' Table_Cleared


    Private Sub Table_NewRow(sender As Object, e As DataTableNewRowEventArgs)
        Console.WriteLine("Table_NewRow Event")

        RecalculateTotal()
    End Sub ' Table_NewRow


    Public Function GetFileName() As String
        Dim dt As String = ""
        dt = System.Reflection.Assembly.GetExecutingAssembly().Location
        dt = System.IO.Path.GetDirectoryName(dt)
        dt = System.IO.Path.Combine(dt, "Dati_Contabilità.xml")

        Return dt
    End Function ' GetFileName


    Public Sub PopulateInitialData(dtThisTable As System.Data.DataTable)
        Dim it As System.Globalization.CultureInfo = New Globalization.CultureInfo("it-CH")
        Dim dr As System.Data.DataRow = Nothing

        For i As Integer = 1 To 12 Step 1
            Dim dt As DateTime = New DateTime(2014, i, 1).AddMonths(1).AddDays(-1)

            dr = dtThisTable.NewRow()

            dr("DATA") = dt
            dr("DETTAGLIATO") = "Mese " + it.TextInfo.ToTitleCase(dt.ToString("MMMM", it))
            dr("ENTRATA") = 123
            dr("USCITA") = 456

            dtThisTable.Rows.Add(dr)
        Next i
    End Sub ' PopulateInitialData



    ' http://stackoverflow.com/questions/24885682/avoid-round-off-decimals-in-datagridview-control
    ' http://stackoverflow.com/questions/11229590/how-to-format-a-column-with-number-decimal-with-max-and-min-in-datagridview
    Public Class RoundedToFiveFormatter
        Implements IFormatProvider
        Implements ICustomFormatter
        Implements IFormattable

        Public Function GetFormat(formatType As Type) As Object Implements IFormatProvider.GetFormat
            If Object.ReferenceEquals(formatType, GetType(ICustomFormatter)) Then
                Return Me
            End If

            Return Nothing
        End Function


        Public Function Format(format1 As String, arg As Object, formatProvider As IFormatProvider) As String Implements ICustomFormatter.Format
            ' Check whether this is an appropriate callback             
            If Not Me.Equals(formatProvider) Then
                Return Nothing
            End If

            If arg Is Nothing Then
                Return Nothing
            End If

            Dim numericString As String = arg.ToString()

            Dim result As Decimal = 0

            If Decimal.TryParse(numericString, result) Then
                Dim it As System.Globalization.CultureInfo = New Globalization.CultureInfo("it-CH")
                Return (Math.Round(result * 20.0, 0, MidpointRounding.AwayFromZero) / 20.0).ToString("N2", it)
            Else
                Return Nothing
            End If
        End Function


        Public Overloads Function ToString(ByVal format As String, ByVal formatProvider As System.IFormatProvider) As String Implements System.IFormattable.ToString
            Return Me.Format(format, Nothing, formatProvider)
        End Function

    End Class




    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If (System.IO.File.Exists(m_FileName)) Then
            m_dt.ReadXml(m_FileName)
            'Else
            ' PopulateInitialData(Me.m_dt)
        End If

        Me.m_bDontSaveWhileLoading = False
        Me.DataGridView1.DataSource = m_dt


        'DataGridView1.RowsDefaultCellStyle.BackColor = Color.LightGray
        DataGridView1.RowsDefaultCellStyle.BackColor = Color.White
        'DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.DarkGray
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.WhiteSmoke

        ' Set the row and column header styles.
        'DataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        'DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black
        'DataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Black



        Me.DataGridView1.Columns("ENTRATA").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView1.Columns("ENTRATA").DefaultCellStyle.Format = "0.00##"
        'Me.DataGridView1.Columns("ENTRATA").DefaultCellStyle.FormatProvider = New RoundedToFiveFormatter()


        Me.DataGridView1.Columns("USCITA").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView1.Columns("USCITA").DefaultCellStyle.Format = "0.00##"

        'Me.DataGridView1.RowHeadersVisible = False

        Me.DataGridView1.Columns("DATA").Width = 80
        Me.DataGridView1.Columns("ENTRATA").Width = 80
        Me.DataGridView1.Columns("USCITA").Width = 80

        Me.DataGridView1.Columns("DETTAGLIATO").Width = Me.DataGridView1.Width - Me.DataGridView1.RowHeadersWidth - Me.DataGridView1.Columns("USCITA").Width - Me.DataGridView1.Columns("ENTRATA").Width - Me.DataGridView1.Columns("DATA").Width - 2

        Me.DataGridView1.Sort(Me.DataGridView1.Columns("DATA"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub ' Form1_Load


    Private Sub RecalculateTotal()
        Dim it As System.Globalization.CultureInfo = New Globalization.CultureInfo("it-CH")

        Dim entrata As Double = 0
        Dim uscita As Double = 0

        For Each dr In m_dt.Rows
            Dim objEntrata As Object = dr("entrata")
            Dim objUscita As Object = dr("uscita")

            If objEntrata IsNot Nothing AndAlso objEntrata IsNot System.DBNull.Value Then
                entrata += dr("entrata")
            End If

            If objUscita IsNot Nothing AndAlso objUscita IsNot System.DBNull.Value Then
                uscita += dr("uscita")
            End If

        Next dr

        Me.lblTotalEntrata.Text = entrata.ToString("N2", it).PadLeft(10, " "c) ' Einnahmen
        Me.lblTotalUscita.Text = uscita.ToString("N2", it).PadLeft(10, " "c) ' Ausgaben

        Dim delta As Double = entrata - uscita
        Me.lblNetto.Text = delta.ToString("N2", it).PadLeft(10, " "c)

        If Not Me.m_bDontSaveWhileLoading Then
            m_dt.WriteXml(m_FileName)
        End If

    End Sub ' RecalculateTotal


End Class
