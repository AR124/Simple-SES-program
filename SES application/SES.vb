Imports System.Data.OleDb
Public Class SES
    'KONEKSI DATABASE----------------------------------------------------------------------------------------
    Public database As New OleDb.OleDbConnection
    Private Constr As String = "Provider=Microsoft.jet.OLEDB.4.0;Data Source=E:\_MATERI_KULIAH\SEMESTER_4\Statistik\Database1.mdb"
    Public Sub koneksi()
        If database.State = ConnectionState.Closed Then
            database.ConnectionString = Constr
            Try
                database.Open()
            Catch ex As Exception
                MsgBox("Koneksi Gagal" & vbCr & ex.ToString)
            End Try
        End If
    End Sub

    'IKI SCRIPT E--------------------------------------------------------------------------------------------
    Sub bersih()
        tb0.Text = 0
        tb1.Text = 0
        tb2.Text = 0
        tb3.Text = 0
        tb4.Text = 0
        tb5.Text = 0
        tb6.Text = 0
        tb7.Text = 0
        tb8.Text = 0
        tb9.Text = 0
        'nilai pemulusan
        npa10.Text = 0
        npa11.Text = 0
        npa12.Text = 0
        npa13.Text = 0
        npa14.Text = 0
        npa15.Text = 0
        npa16.Text = 0
        npa17.Text = 0
        npa18.Text = 0
        npa19.Text = 0
        npa50.Text = 0
        npa51.Text = 0
        npa52.Text = 0
        npa53.Text = 0
        npa54.Text = 0
        npa55.Text = 0
        npa56.Text = 0
        npa57.Text = 0
        npa58.Text = 0
        npa59.Text = 0
        npa90.Text = 0
        npa91.Text = 0
        npa92.Text = 0
        npa93.Text = 0
        npa94.Text = 0
        npa95.Text = 0
        npa96.Text = 0
        npa97.Text = 0
        npa98.Text = 0
        npa99.Text = 0
        'nilai kesalahan
        nka10.Text = 0
        nka11.Text = 0
        nka12.Text = 0
        nka13.Text = 0
        nka14.Text = 0
        nka15.Text = 0
        nka16.Text = 0
        nka17.Text = 0
        nka18.Text = 0
        nka19.Text = 0
        nka50.Text = 0
        nka51.Text = 0
        nka52.Text = 0
        nka53.Text = 0
        nka54.Text = 0
        nka55.Text = 0
        nka56.Text = 0
        nka57.Text = 0
        nka58.Text = 0
        nka59.Text = 0
        nka90.Text = 0
        nka91.Text = 0
        nka92.Text = 0
        nka93.Text = 0
        nka94.Text = 0
        nka95.Text = 0
        nka96.Text = 0
        nka97.Text = 0
        nka98.Text = 0
        nka99.Text = 0
        'kuadrat
        knka10.Text = 0
        knka11.Text = 0
        knka12.Text = 0
        knka13.Text = 0
        knka14.Text = 0
        knka15.Text = 0
        knka16.Text = 0
        knka17.Text = 0
        knka18.Text = 0
        knka19.Text = 0
        knka50.Text = 0
        knka51.Text = 0
        knka52.Text = 0
        knka53.Text = 0
        knka54.Text = 0
        knka55.Text = 0
        knka56.Text = 0
        knka57.Text = 0
        knka58.Text = 0
        knka59.Text = 0
        knka90.Text = 0
        knka91.Text = 0
        knka92.Text = 0
        knka93.Text = 0
        knka94.Text = 0
        knka95.Text = 0
        knka96.Text = 0
        knka97.Text = 0
        knka98.Text = 0
        knka99.Text = 0
        'jumlah
        tjumlaha1.Text = 0
        tjumlaha5.Text = 0
        tjumlaha9.Text = 0
        'nilai mse
        tmsea1.Text = 0
        tmsea5.Text = 0
        tmsea9.Text = 0
    End Sub

    Public alfa1 As Double = 0.1
    Public alfa2 As Double = 0.5
    Public alfa3 As Double = 0.9

    Private Sub bthitung_Click(sender As Object, e As EventArgs) Handles bthitung.Click
        'nilai pengamatan---------------------------------------------------------------------
        '1
        npa11.Text = alfa1 * tb1.Text + (1 - alfa1) * tb1.Text
        npa51.Text = alfa2 * tb1.Text + (1 - alfa2) * tb1.Text
        npa91.Text = alfa3 * tb1.Text + (1 - alfa3) * tb1.Text
        '2
        npa92.Text = alfa3 * npa91.Text + (1 - alfa3) * tb1.Text
        npa52.Text = alfa2 * npa51.Text + (1 - alfa2) * tb1.Text
        npa12.Text = alfa1 * npa11.Text + (1 - alfa1) * tb1.Text
        '3
        npa93.Text = alfa3 * npa92.Text + (1 - alfa3) * tb2.Text
        npa53.Text = alfa2 * npa52.Text + (1 - alfa2) * tb2.Text
        npa13.Text = alfa1 * npa12.Text + (1 - alfa1) * tb2.Text
        '4
        npa94.Text = alfa3 * npa93.Text + (1 - alfa3) * tb3.Text
        npa54.Text = alfa2 * npa53.Text + (1 - alfa2) * tb3.Text
        npa14.Text = alfa1 * npa13.Text + (1 - alfa1) * tb3.Text
        '5
        npa95.Text = alfa3 * npa94.Text + (1 - alfa3) * tb4.Text
        npa55.Text = alfa2 * npa54.Text + (1 - alfa2) * tb4.Text
        npa15.Text = alfa1 * npa14.Text + (1 - alfa1) * tb4.Text
        '6
        npa96.Text = alfa3 * npa95.Text + (1 - alfa3) * tb5.Text
        npa56.Text = alfa2 * npa55.Text + (1 - alfa2) * tb5.Text
        npa16.Text = alfa1 * npa15.Text + (1 - alfa1) * tb5.Text
        '7
        npa97.Text = alfa3 * npa96.Text + (1 - alfa3) * tb6.Text
        npa57.Text = alfa2 * npa56.Text + (1 - alfa2) * tb6.Text
        npa17.Text = alfa1 * npa16.Text + (1 - alfa1) * tb6.Text
        '8
        npa98.Text = alfa3 * npa97.Text + (1 - alfa3) * tb7.Text
        npa58.Text = alfa2 * npa57.Text + (1 - alfa2) * tb7.Text
        npa18.Text = alfa1 * npa17.Text + (1 - alfa1) * tb7.Text
        '9
        npa99.Text = alfa3 * npa98.Text + (1 - alfa3) * tb8.Text
        npa59.Text = alfa2 * npa58.Text + (1 - alfa2) * tb8.Text
        npa19.Text = alfa1 * npa18.Text + (1 - alfa1) * tb8.Text
        '10
        npa90.Text = alfa3 * npa99.Text + (1 - alfa3) * tb9.Text
        npa50.Text = alfa2 * npa59.Text + (1 - alfa2) * tb9.Text
        npa10.Text = alfa1 * npa19.Text + (1 - alfa1) * tb9.Text

        'nilai kesalahan--------------------------------------------------------------------------------
        'alfa=0.1
        nka11.Text = tb1.Text - npa11.Text
        nka12.Text = tb2.Text - npa12.Text
        nka13.Text = tb3.Text - npa13.Text
        nka14.Text = tb4.Text - npa14.Text
        nka15.Text = tb5.Text - npa15.Text
        nka16.Text = tb6.Text - npa16.Text
        nka17.Text = tb7.Text - npa17.Text
        nka18.Text = tb8.Text - npa18.Text
        nka19.Text = tb9.Text - npa19.Text
        nka10.Text = tb0.Text - npa10.Text
        'alfa=0.5
        nka51.Text = tb1.Text - npa51.Text
        nka52.Text = tb2.Text - npa52.Text
        nka53.Text = tb3.Text - npa53.Text
        nka54.Text = tb4.Text - npa54.Text
        nka55.Text = tb5.Text - npa55.Text
        nka56.Text = tb6.Text - npa56.Text
        nka57.Text = tb7.Text - npa57.Text
        nka58.Text = tb8.Text - npa58.Text
        nka59.Text = tb9.Text - npa59.Text
        nka50.Text = tb0.Text - npa50.Text
        'alfa=0.9
        nka91.Text = tb1.Text - npa91.Text
        nka92.Text = tb2.Text - npa92.Text
        nka93.Text = tb3.Text - npa93.Text
        nka94.Text = tb4.Text - npa94.Text
        nka95.Text = tb5.Text - npa95.Text
        nka96.Text = tb6.Text - npa96.Text
        nka97.Text = tb7.Text - npa97.Text
        nka98.Text = tb8.Text - npa98.Text
        nka99.Text = tb9.Text - npa99.Text
        nka90.Text = tb0.Text - npa90.Text

        'kuadrat nilai kesalahan-------------------------------------------------------------------------------------------
        'alfa=0.1
        knka11.Text = nka11.Text ^ 2
        knka12.Text = nka12.Text ^ 2
        knka13.Text = nka13.Text ^ 2
        knka14.Text = nka14.Text ^ 2
        knka15.Text = nka15.Text ^ 2
        knka16.Text = nka16.Text ^ 2
        knka17.Text = nka17.Text ^ 2
        knka18.Text = nka18.Text ^ 2
        knka19.Text = nka19.Text ^ 2
        knka10.Text = nka10.Text ^ 2
        'alfa=0.5
        knka51.Text = nka51.Text ^ 2
        knka52.Text = nka52.Text ^ 2
        knka53.Text = nka53.Text ^ 2
        knka54.Text = nka54.Text ^ 2
        knka55.Text = nka55.Text ^ 2
        knka56.Text = nka56.Text ^ 2
        knka57.Text = nka57.Text ^ 2
        knka58.Text = nka58.Text ^ 2
        knka59.Text = nka59.Text ^ 2
        knka50.Text = nka50.Text ^ 2
        'alfa=0.9
        knka91.Text = nka91.Text ^ 2
        knka92.Text = nka92.Text ^ 2
        knka93.Text = nka93.Text ^ 2
        knka94.Text = nka94.Text ^ 2
        knka95.Text = nka95.Text ^ 2
        knka96.Text = nka96.Text ^ 2
        knka97.Text = nka97.Text ^ 2
        knka98.Text = nka98.Text ^ 2
        knka99.Text = nka99.Text ^ 2
        knka90.Text = nka90.Text ^ 2

        'deklarasi variabel jumlah a=0,1
        Dim h1, h2, h3, h4, h5, h6, h7, h8, h9, h10 As Double
        h1 = knka10.Text
        h2 = knka11.Text
        h3 = knka12.Text
        h4 = knka13.Text
        h5 = knka14.Text
        h6 = knka15.Text
        h7 = knka16.Text
        h8 = knka17.Text
        h9 = knka18.Text
        h10 = knka19.Text

        'deklarasi variabel jumlah a=0,5
        Dim i1, i2, i3, i4, i5, i6, i7, i8, i9, i10 As Double
        i1 = knka50.Text
        i2 = knka51.Text
        i3 = knka52.Text
        i4 = knka53.Text
        i5 = knka54.Text
        i6 = knka55.Text
        i7 = knka56.Text
        i8 = knka57.Text
        i9 = knka58.Text
        i10 = knka59.Text

        'deklarasi variabel jumlah a=0,9
        Dim j1, j2, j3, j4, j5, j6, j7, j8, j9, j10 As Double
        j1 = knka90.Text
        j2 = knka91.Text
        j3 = knka92.Text
        j4 = knka93.Text
        j5 = knka94.Text
        j6 = knka95.Text
        j7 = knka96.Text
        j8 = knka97.Text
        j9 = knka98.Text
        j10 = knka99.Text

        'jumlah
        tjumlaha1.Text = h1 + h2 + h3 + h4 + h5 + h6 + h7 + h8 + h9 + h10
        tjumlaha5.Text = i1 + i2 + i3 + i4 + i5 + i6 + i7 + i8 + i9 + i10
        tjumlaha9.Text = j1 + j2 + j3 + j4 + j5 + j6 + j7 + j8 + j9 + j10

        'nilai mse
        tmsea1.Text = tjumlaha1.Text * 10
        tmsea5.Text = tjumlaha5.Text * 10
        tmsea9.Text = tjumlaha9.Text * 10
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call bersih()
    End Sub

End Class
