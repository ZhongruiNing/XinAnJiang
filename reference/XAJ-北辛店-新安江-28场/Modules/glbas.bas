Attribute VB_Name = "glbas"
Option Explicit
Global Path As String, pathc As String, pathd As String, pathp As String, paths As String, pathy As String
Global gltt As Integer, glh As String, glhh As Long '次洪时段长
Global NoFlood As Integer    '年号
Global NumberNo As Integer   '年数
Global FloodNo() As Long  '年号数组
Global NoYear As Long, NoYearr As Integer '年
Global StartTimeD As Long, EndTimeD As Long, LongTimeD As Integer
Global StartTime As Long, EndTime As Long, LongTime As Integer
Global TimeStart As Date, TimeEnd As Date
Global gsdsj() As Long, sdsj() As Long, Na As Integer, Ma As Integer, na1 As Integer, para4() As Single, glchsdsj() As Long, xq As Integer
Global glnn As Integer, glnn1 As Integer, glday As Integer, glmm As Integer, gltm As Integer
Global glchmn() As Date, glchmm() As Date, glchml() As Long, Ddata() As Date, CountFlood As Integer
Global sdsd() As Date, sdsw() As Date
Global para1(23) As Single, para2() As Single, para3() As String
Global BeginDay As Integer
Global Hdate() As Date, UU As String, PK As Single
Global Qjliu As Single, Q00() As Single, Qjiliu As Single, Qmaxx() As Single, Unitt As String, Unit() As String, Showw As String, jy As Integer
Global glpwhf() As Single, glqobshf() As Single, glqcalhf() As Single, glqadjhf() As Single
Global BasinCName() As String, BasineEName() As String, Basin As String
Global BasinNa As Integer, rc As Integer, kp As Single
Global dyly As String, dylyc As String, Zhanevap As String
Global woto As Single, wcto As Single
Global glztz As Integer, DAindex As Integer
Global qsig() As Single
Global kname As String, bname As String, sql1 As String
Global cn  As New ADODB.Connection
Global cmd As New ADODB.Command
Global b As New ADODB.Recordset
Global rd7 As New ADODB.Recordset
Global rd8 As New ADODB.Recordset
Global rst As New ADODB.Recordset
Global rdd As New ADODB.Recordset
Global rd1 As New ADODB.Recordset
Global bb As New ADODB.Recordset
Global rd2 As New ADODB.Recordset
Sub Main()
ReDim Number(2), NoStation(2), BasinCName(2), BasineEName(2)
Dim i As Integer
Path = App.Path
pathc = Path + "\chdat\"
kname = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathc + "data.mdb;Persist Security Info=False"
cn.Open kname
Set cmd.ActiveConnection = cn

Showw = "MX"
BeginDay = 30

gltt = 1
gltm = 6#

Dim k1 As Single, k2 As Single
k1 = 1 / Log(0.9983)
k2 = 1 / Log(0.9974)
MDImain.Show

End Sub

