'��ģ����Ҫ�����������ִ��������,��Ҫ��������



public K1Array   '''����˵�1 ��K������Ĺ�Ʊ
public K12Array  '''����˵�1��2��K������Ĺ�Ʊ
public K13Array  '''����˵�1��3��K������Ĺ�Ʊ
public K1Market
public K12Market
public K13Market


public K123Array '''��3��K�߳����Ժ���'
public K134Array '''��4��K�߳����Ժ���
public K124Array '''��4��K�߳����Ժ���

public K123Market
public K1342Market
public K1243Market





'''Set strNotice = "[ ��ʾ: ]"
'''Set strWarning= "[ ����: ]"
'''Set strError  = "[ ����: ]"


public DBConnection		'''���ݿ�����'''

public SYSOpenRate		'''���ַ���'''


Set K1Array = CreateObject("Stock.ArrayString")
Set K12Array = CreateObject("Stock.ArrayString")
Set K13Array = CreateObject("Stock.ArrayString")
Set K123Array = CreateObject("Stock.ArrayString")
Set K134Array = CreateObject("Stock.ArrayString")
Set K124Array = CreateObject("Stock.ArrayString")


Set K1Market = CreateObject("Stock.ArrayString")
Set K12Market = CreateObject("Stock.ArrayString")
Set K13Market = CreateObject("Stock.ArrayString")
Set K123Market = CreateObject("Stock.ArrayString")
Set K134Market = CreateObject("Stock.ArrayString")
Set K124Market = CreateObject("Stock.ArrayString")


'''������ʱ�������칺��Ĺ�Ʊ���۸�ʹ��룬���ڼ�������Ͳ�������
Set HistoryArray = CreateObject("Stock.ArrayString")
Set HistoryPrice = CreateObject("Stock.Array")
Set HistoryCondition = CreateObject("Stock.ArrayString")



'''ע���һ��K������9.45���γɵ�
Function todayOpenTime()
    Dim t 
    t = Date
    todayOpenTime = Year(t) & "/" & _
    Right("0" & Month(t),2)  & "/" & _
    Right("0" & Day(t),2)  & " " & _  
    Right("0" & "09",2) & ":" &_
    Right("0" & "45",2) & ":" &_
    Right("0" & "00",2)
End Function


Function getTime(h,m,s)
    Dim t 
    t = Date
    getTime = Year(t) & "/" & _
    Right("0" & Month(t),2)  & "/" & _
    Right("0" & Day(t),2)  & " " & _  
    Right("0" & h,2) & ":" &_
    Right("0" & m,2) & ":" &_
    Right("0" & s,2)
End Function



'''����ѡ��ʱ��־���
Function debugString(String)
	call Document.DebugFile("C:\Weisoft Stock(x64)\Log\AutoTrader.TXT",String,1)
End Function

Function logString(String)
	call Document.DebugFile("C:\Weisoft Stock(x64)\Log\LogAutoTrader.TXT",String,1)
End Function

'''��1�׶Σ�ֻ�ǽ�123�����ɾ��
Function removeStage1()
	K123Array.RemoveAll()
	K123Market.RemoveAll()
End Function

'''��2�׶���ɺ�����ѡ�����Ĺ�Ʊ��Ӧ�������� 
Function removeStage2()
	K1Array.RemoveAll()
	K12Array.RemoveAll()
	K13Array.RemoveAll()
	K1Market.RemoveAll()
	K12Market.RemoveAll()
	K13Market.RemoveAll()
	K124Array.RemoveAll()
	K134Array.RemoveAll()
	K124Market.RemoveAll()
	K134Market.RemoveAll()
End Function

'''��3�׶���ɺ���������Ĺ�Ʊ�Ѿ�ƽ����
Function removeStage3()
	HistoryArray.RemoveAll()
	HistoryPrice.RemoveAll()
	HistoryCondition.RemoveAll()
End Function


Function getOpenRate()
	Dim rate
	strSQL = "SELECT * FROM sys_rate"
	Set rs = CreateObject("ADODB.Recordset")
	rs.Open strSQL, DBConnection, , , adCmdText
	rate = rs.Fields("open_rate")
	rs.Close
	getOpenRate = rate
End Function

'''�ú������ڼ��ִ�л�������ִ�л����Ƿ�׼��ִ��
Function envirnmentCheck()

	'''�˻��Ƿ��¼
	if Order.IsAccount("6004625") = 0 then
		debugString("ע��Ҫ�����ˣ����˻���û�е�¼�����¼")
	end if
	'''�ʽ��Ƿ��㹻
	
	''''
	
	
	
	'''��ȡ����
	openConnection()
	SYSOpenRate = getOpenRate()
	closeConnection()	
	
	logString("[ ��ʾ: ] ���ַ���:" & SYSOpenRate)
	
	 


End Function


'''K1 �źŷ�����
Function K1Signal(Label,Market)
	Dim BarBegin
	Dim BarEnd
	Dim CaculateEnd
	Dim r
	r = 0
	Set Formula = marketdata.STKINDI(Label,Market,"morningk1",2,2)
	BarBegin = Formula.GetPosFromDate(todayOpenTime)
	BarEnd = Formula.DataSize-1
	CaculateEnd = BarBegin+0
	for i=BarBegin to CaculateEnd
		if Formula.GetBufData("Output",i) = 1 then
			r = 1
			exit for
		end if
	next
	K1Signal = r
End Function


'''K1 and K2 �źŷ�����
Function K12Signal(Label,Market)
	Dim BarBegin
	Dim BarEnd
	Dim CaculateEnd
	Dim r
	r = 0
	Set Formula = marketdata.STKINDI(Label,Market,"morningk12",2,2)
	BarBegin = Formula.GetPosFromDate(todayOpenTime)
	CaculateEnd = BarBegin+1
	for j=BarBegin to CaculateEnd
		if Formula.GetBufData("Output",j) = 1 then
			r = 1
			exit for
		end if
	next
	K12Signal = r
End Function



'''K1 and K3 �źŷ�����
Function K13Signal(Label,Market)
	Dim BarBegin
	Dim BarEnd
	Dim CaculateEnd
	Dim r
	r = 0
	Set Formula = marketdata.STKINDI(Label,Market,"morningk13",2,2)
	BarBegin = Formula.GetPosFromDate(todayOpenTime)
	CaculateEnd = BarBegin+2
	for j=BarBegin to CaculateEnd
		if Formula.GetBufData("Output",j) = 1 then
			r = 1
			exit for
		end if
	next
	K13Signal = r
End Function



'''K1 and K2 and  K3 �źŷ�����
Function K123Signal(Label,Market)
	Dim BarBegin
	Dim BarEnd
	Dim CaculateEnd
	Dim r
	r = 0
	Set Formula = marketdata.STKINDI(Label,Market,"morningk123",2,2)
	BarBegin = Formula.GetPosFromDate(todayOpenTime)
	CaculateEnd = BarBegin+2
	for j=BarBegin to CaculateEnd
		if Formula.GetBufData("Output",j) = 1 then
			r = 1
			exit for
		end if
	next
	K123Signal = r
End Function

'''K1 and K2 and  K4 �źŷ�����
Function K124Signal(Label,Market)
	Dim BarBegin
	Dim BarEnd
	Dim CaculateEnd
	Dim r
	r = 0
	Set Formula = marketdata.STKINDI(Label,Market,"morningk124",2,2)
	BarBegin = Formula.GetPosFromDate(todayOpenTime)
	CaculateEnd = BarBegin+3
	for j=BarBegin to CaculateEnd
		if Formula.GetBufData("Output",j) = 1 then
			r = 1
			exit for
		end if
	next
	K124Signal = r
End Function


'''K1 and K3 and  K4 �źŷ�����
Function K134Signal(Label,Market)
	Dim BarBegin
	Dim BarEnd
	Dim CaculateEnd
	Dim r
	r = 0
	Set Formula = marketdata.STKINDI(Label,Market,"morningk134",2,2)
	BarBegin = Formula.GetPosFromDate(todayOpenTime)
	CaculateEnd = BarBegin+3
	for j=BarBegin to CaculateEnd
		if Formula.GetBufData("Output",j) = 1 then
			r = 1
			exit for
		end if
	next
	K134Signal = r
End Function



'''����������K1�����Ĺ�Ʊѡ�ٳ���,��һ����ʱ��Ҫ���Ǻܸߣ���������Application.PeekAndPump 
Function runK1()
	Count = marketdata.GetReportCount("SZ")
	for i=0 to Count-1
		Set Report1  = marketdata.GetReportDataByIndex("SZ",i)
		if Left(Report1.Label,3) = "002" or Left(Report1.Label,3) = "300"  or Left(Report1.Label,3) = "000" or Left(Report1.Label,3) = "600" or Left(Report1.Label,3) = "601"  or Left(Report1.Label,3) = "603" then
			if K1Signal(Report1.Label,"SZ")=1 then
				K1Array.AddBack(Report1.Label)
				K1Market.AddBack(Report1.MarketName)
			end if
		end if 
		Application.PeekAndPump 
	next
	Count = marketdata.GetReportCount("SH")
	for i=0 to Count-1
		Set Report1  = marketdata.GetReportDataByIndex("SH",i)
		if Left(Report1.Label,3) = "600" or Left(Report1.Label,3) = "601"  or Left(Report1.Label,3) = "603" or Left(Report1.Label,3) = "002" or Left(Report1.Label,3) = "300"  or Left(Report1.Label,3) = "000" then
			if K1Signal(Report1.Label,"SH")=1 then
				K1Array.AddBack(Report1.Label)
				K1Market.AddBack(Report1.MarketName)
			end if
		end if
		Application.PeekAndPump 
	next
	runK1 = K1Array.Count
End Function



'''����������K1 and K2�����Ĺ�Ʊѡ�ٳ���,�ڶ�����ʱ��Ҫ���Ǻܸߣ���������Application.PeekAndPump 
Function runK12()
	Count = K1Array.Count
	for i=0 to Count-1
		Code = K1Array.GetAt(i)
		Market = K1Market.GetAt(i)
		Set Report1  = marketdata.GetReportData(Code,Market)
		if K12Signal(Code,Market)=1 then
			K12Array.AddBack(Report1.Label)
			K12Market.AddBack(Report1.MarketName)
		end if
		call Application.PeekAndPump
	next
	runK12 = K12Array.Count
End Function

'''����������K1 and K3�����Ĺ�Ʊѡ�ٳ���
Function runK13()
	Count = K1Array.Count
	for i=0 to Count-1
		Code = K1Array.GetAt(i)
		Market = K1Market.GetAt(i)
		Set Report1  = marketdata.GetReportData(Code,Market)
		if K13Signal(Code,Market)=1 then
			K13Array.AddBack(Report1.Label)
			K13Market.AddBack(Report1.MarketName)
		end if
		call Application.PeekAndPump
	next
	runK13 = K13Array.Count
End Function

'''����������K1 and K2 and K3�����Ĺ�Ʊѡ�ٳ���
Function runK123()
	Count = K12Array.Count
	for i=0 to Count-1
		Code = K12Array.GetAt(i)
		Market = K12Market.GetAt(i)
		Set Report1  = marketdata.GetReportData(Code,Market)
		if K123Signal(Code,Market)=1 then
			K123Array.AddBack(Code)
			K123Market.AddBack(Market)
			openConnection()
			insertConditionTable Code,Market,"123",Time,Date
			closeConnection()
		end if
	next
	runK123 = K123Array.Count
End Function


Function exclude123(Code)
	Dim r
	r = 1
	Count = K123Array.Count
	for i=0 to Count-1
		Level = K123Array.GetAt(i)
		if Level = Code then
			r=0
			exit for
		end if
	next
	exclude123=r
End Function



Function exclude124(Code)
	Dim r
	r = 1
	Count = K124Array.Count
	for i=0 to Count-1
		Level = K124Array.GetAt(i)
		if Level = Code then
			r=0
			exit for
		end if
	next
	exclude124=r
End Function


'''����������K1 and K2 and K4�����Ĺ�Ʊѡ�ٳ���
Function runK124()
	Count = K12Array.Count
	for i=0 to Count-1
		Code = K12Array.GetAt(i)
		Market = K12Market.GetAt(i)
		Set Report1  = marketdata.GetReportData(Code,Market)
		if K124Signal(Code,Market)=1 and exclude123(Code)=1 then  '''�����ų�123�ظ���
			K124Array.AddBack(Code)
			K124Market.AddBack(Market)
			openConnection()
			insertConditionTable Code,Market,"124",Time,Date
			closeConnection()
		end if
	next
	runK124 = K124Array.Count
End Function

'''����������K1 and K3 and K4�����Ĺ�Ʊѡ�ٳ���
Function runK134()
	Count = K13Array.Count
	for i=0 to Count-1
		Code = K13Array.GetAt(i)
		Market = K13Market.GetAt(i)
		Set Report1  = marketdata.GetReportData(Code,Market)
		if K134Signal(Code,Market)=1 and exclude124(Code)=1 then   '''�����ų�124�ظ���
			K134Array.AddBack(Code)
			K134Market.AddBack(Market)
			openConnection()
			insertConditionTable Code,Market,"134",Time,Date
			closeConnection()
		end if
	next
	runK134 = K134Array.Count
End Function



'''��һ�ι���Ĳ���
Function firstBuy()

	Dim expendable_cash		'''��ǰ�����ʽ�
	Dim each_stock_cash		'''ÿֻ��Ʊ�����ʽ�
	Dim count 				'''�м�ֻ��Ʊ
	
	Dim	stock_count			'''ÿֻ��Ʊ����Ĺ�Ʊ����
	
	
	count  = K123Array.Count
	if count>0 then			'''�й�Ʊ����Ž��в���
		expendable_cash = Order.Account2(3,"")
		expendable_cash = expendable_cash / 2   '''����տ�ʼֻ�������ʽ��1�룬��������Զ�����
		
		expendable_cash = expendable_cash*(1-SYSOpenRate)   '''�ֳ�һ���ȥ��������
		each_stock_cash = expendable_cash / count  '''����ÿֻ��Ʊ�����ʽ�
		
		for i=0 to count-1
			Code = K123Array.GetAt(i)
			Market = K123Market.GetAt(i)
			set Report1 = marketdata.GetReportData(Code,Market)
			stock_count  = each_stock_cash / (Report1.NewPrice)  '''����Ҫ���Ϸ���
			stock_count = Int(Int(stock_count)/100)*100    '''ת���ɱ�׼����
			
			rs = Order.Buy(1,stock_count,Report1.NewPrice,0,Code,Market,"",0)
			if rs <> -1 then
				logString("[ ��ʾ: ]" & Code & " �Լ۸�:" & Report1.NewPrice & " ���룺" & stock_count & "��")
			else
				logString("[ ����: ] ���� " & Code & " ʧ��")
			end if
		next
	end if
	

End Function


Function emailSender(Subs,Msg)
    Set mail = CreateObject("WWSCommon.SmtpMail")
    with mail
         .SenderName = "duanwujie"
         .SenderAddress = "this is email address"
         .Subject = Subs
     end with
     call mail.AddReceiver("test","this is email addres")
     call mail.AddTextContent(Msg)
     call mail.Sender("smtp.163.com","this is email addres","this is passsword")
End Function


'''�ڶ��ι���Ĳ���
Function secondBuy()

	Dim expendable_cash		'''��ǰ�����ʽ�
	Dim each_stock_cash		'''ÿֻ��Ʊ�����ʽ�
	Dim count1 				'''124�м�ֻ��Ʊ
	Dim count2              '''134�м�ֻ��Ʊ
	Dim count				'''124+134�м�ֻ��Ʊ
	Dim	stock_count			'''ÿֻ��Ʊ����Ĺ�Ʊ����
	
	count1  = K124Array.Count
	count2  = K134Array.Count
	count = count1+count2
	
	debugString(count & ":" & count1 & ":" & count2)
	
	if count>0 then			'''�й�Ʊ����Ž��в���
		logString("[ ��ʾ: ] ��4��K����" & count & "ֻ��Ʊ����")
		expendable_cash = Order.Account2(3,"")
		
		expendable_cash = expendable_cash*(1-SYSOpenRate)   '''ȥ��������
		
		expendable_cash = expendable_cash  '''�������ʽ����ڹ����Ʊ
		each_stock_cash = expendable_cash / count  '''����ÿֻ��Ʊ�����ʽ�
		logString("each_stock_cash:" & each_stock_cash & " openrate:" & SYSOpenRate) 
	else
		logString("[ ��ʾ: ] ��4��K��û�й�Ʊ����")
	end if
	
	
	
	
	'''��������124�����Ĺ�Ʊ
	for i=0 to count1-1
		Code = K124Array.GetAt(i)
		Market = K124Market.GetAt(i)
		set Report1 = marketdata.GetReportData(Code,Market)
		stock_count1  = each_stock_cash / Report1.NewPrice
		stock_count1 = Int(Int(stock_count1)/100)*100    '''ת���ɱ�׼����
		
		debugString("each_stock_cash:" & each_stock_cash  & "price:" & Report1.NewPrice & "stock_count1:" & stock_count1)

		rs = Order.Buy(1,stock_count1,Report1.NewPrice,0,Code,Market,"",0)
		if rs  <> -1 then
			logString("[ ��ʾ: ]" & Code & " �Լ۸�:" & Report1.NewPrice & " ���룺" & stock_count1 & "��")
		else
			logString("[ ����: ] ���� " & Code & " ʧ��")
		end if
	next
	
	'''��������134�����Ĺ�Ʊ
	for i=0 to count2-1
		Code = K134Array.GetAt(i)
		Market = K134Market.GetAt(i)
		set Report1 = marketdata.GetReportData(Code,Market)
		stock_count2  = each_stock_cash / Report1.NewPrice
		stock_count2 = Int(Int(stock_count2)/100)*100    '''ת���ɱ�׼����
		rs = Order.Buy(1,stock_count2,Report1.NewPrice,0,Code,Market,"",0)
		if rs <> -1 then
			logString("[ ��ʾ: ]" & Code & " �Լ۸�:" & Report1.NewPrice & " ���룺" & stock_count2 & "��")
		else
			logString("[ ����: ] ���� " & Code & " ʧ��")
		end if
	next

End Function


Function getCondtion(Code)
	Dim cond
	strSQL = "SELECT * FROM tmp_condition where code= '"& Code &"'" 
	Set rs = CreateObject("ADODB.Recordset")
	rs.Open strSQL, DBConnection, , , adCmdText
	cond = rs.Fields("cond")
	rs.Close
	getCondtion = cond
End Function


'''���ｫ���칺��Ĺ�Ʊ���Կ��̺�ļ۸�ƽ��
Function closeYestodayOrder()
	removeStage3()			'''��stage3��״̬��ʼ��
	openConnection()		'''�����ݿ�
	strSQL = "SELECT * FROM opened"
	Set rs = CreateObject("ADODB.Recordset")
	rs.Open strSQL, DBConnection, , , adCmdText
	With rs
		if .EOF then
			debugString("[ ���� ]:����Ŀ��̱���û�����ݣ�����޷������ݿ��ж�ȡ����������ƽ��")
			debugString("[ ���� ]:�����н����˻��ж�ȡ��������������ƽ��")
			emailSender "[ ���� ]","����Ŀ��̱���û�����ݣ�����޷������ݿ��ж�ȡ����������ƽ��,���ֶ�ƽ��"
		end if
		
		While (Not .EOF)
    		Label = rs.Fields("code")
    		Market= rs.Fields("market")
    		Vol	  = rs.Fields("lots")
    		Price = rs.Fields("price")
			Cond  = getCondtion(Label)
    		HistoryArray.AddBack(Label)
    		HistoryPrice.AddBack(Price)
			HistoryCondition.AddBack(Cond)
    		set Report1 = marketdata.GetReportData(Label,Market)
			call Order.Sell(1,Vol,Report1.NewPrice,0,Label,Market,"",0)
    		.MoveNext
		Wend
		.Close
	End With
	closeConnection()		'''�ر����ݿ�
End Function




Function clearOpenedTable()
	openConnection()		'''�����ݿ�
	strSQL ="DELETE * from opened"
	DBConnection.Execute strSQL
	closeConnection()		'''�ر����ݿ�
End Function


Function clearTmpconditionTable()
	openConnection()		'''�����ݿ�
	strSQL ="DELETE * from tmp_condition"
	DBConnection.Execute strSQL
	closeConnection()		'''�ر����ݿ�
End Function


Function openConnection()
	set DBConnection=CreateObject("ADODB.connection")
	DBConnection.Provider="Microsoft.ACE.OLEDB.12.0"
	DBConnection.Open "C:\Weisoft Stock(x64)\traderinfo.mdb"
End Function


Function closeConnection()
	DBConnection.Close
End Function


Function insertOpenedTable(sCode,sMarket,sPrice,sLots,sTime,sDate)
	strSQL = "INSERT INTO opened (code,market,price,lots,times,dates) VALUES ('" & sCode & "','" & sMarket & "',"& sPrice &"," & sLots & ", '"& sTime &"','"& sDate &"')"
	DBConnection.Execute strSQL
End Function


Function insertClosedTable(sCode,sMarket,sOpen,sClose,sLots,sProfit,sTime,sDate,sCond)
	strSQL = "INSERT INTO closed (code,market,open_price,close_price,lots,profit,times,dates,cond) VALUES ('" & sCode & "','" & sMarket & "',"& sOpen &","& sClose &"," & sLots & ", "& sProfit &", '"& sTime &"','"& sDate &"','"& sCond &"')"
	DBConnection.Execute strSQL
End Function

Function insertConditionTable(sCode,sMarket,sCond,sTime,sDate)
	strSQL = "INSERT INTO tmp_condition (code,market,cond,times,dates) VALUES ('" & sCode & "','" & sMarket & "','"& sCond &"','"& sTime &"','"& sDate &"')"
	DBConnection.Execute strSQL
End Function



Sub K0()
	envirnmentCheck()
End Sub

'''���ڵ���1����
Sub K1()
	Count = runK1()
	debugString("�Ѿ�ɸѡ������Ϊ1�Ĺ�Ʊ:" & Count & "ֻ")
End Sub

'''���ڵ���12����
Sub K2()
	Count = runK12()
	debugString("�Ѿ�ɸѡ������Ϊ12�Ĺ�Ʊ:" & Count & "ֻ")
End Sub


'''���ڵ���13, 123����
Sub K3()
	Count = runK123()
	debugString("�Ѿ�ɸѡ������Ϊ123�Ĺ�Ʊ:" & Count & "ֻ")
	Count = runK13()
	debugString("�Ѿ�ɸѡ������Ϊ13�Ĺ�Ʊ:" & Count & "ֻ")
End Sub

'''���ڵ���124, 134����
Sub K4()
	Count = runK124()
	debugString("�Ѿ�ɸѡ������Ϊ124�Ĺ�Ʊ:" & Count & "ֻ")
	Count = runK134()
	debugString("�Ѿ�ɸѡ������Ϊ134�Ĺ�Ʊ:" & Count & "ֻ")
End Sub



'''���ڲ��Ե�1�׶εĹ���
Sub B1()
	firstBuy()
	removeStage1()
End Sub

'''���ڲ��Ե�2�׶εĹ���
Sub B2()
	K124Array.AddBack("000425")
	K124Market.AddBack("SZ")
	
	openConnection()
	SYSOpenRate = getOpenRate()
	closeConnection()
	
	secondBuy()
	removeStage2()
End Sub

'''���ڲ�����ʷ���ݣ����в���,��opened,tmp_condition���в�������
Sub TestBuy()
	clearOpenedTable()
	clearTmpconditionTable()
	openConnection()
	insertOpenedTable "000001","SH",1234556,1000,Time,Date
	insertOpenedTable "000002","SH",1234556,2000,Time,Date
	insertOpenedTable "000003","SH",1234556,3000,Time,Date
	insertOpenedTable "000004","SH",1234556,4000,Time,Date
	insertOpenedTable "000005","SH",1234556,5000,Time,Date
	insertOpenedTable "000006","SH",1234556,6000,Time,Date
	insertOpenedTable "000007","SH",1234556,7000,Time,Date
	insertOpenedTable "000008","SH",1234556,8000,Time,Date
	insertOpenedTable "000009","SH",1234556,9000,Time,Date
	insertOpenedTable "000010","SH",1234556,11000,Time,Date
	insertOpenedTable "000011","SH",1234556,12000,Time,Date
	insertOpenedTable "000012","SH",1234556,13000,Time,Date

	
	insertConditionTable "000001","SH","123",Time,Date
	insertConditionTable "000002","SH","123",Time,Date
	insertConditionTable "000003","SH","123",Time,Date
	insertConditionTable "000004","SH","123",Time,Date
	insertConditionTable "000005","SH","123",Time,Date
	insertConditionTable "000006","SH","124",Time,Date
	insertConditionTable "000007","SH","124",Time,Date
	insertConditionTable "000008","SH","124",Time,Date
	insertConditionTable "000009","SH","124",Time,Date
	insertConditionTable "000010","SH","134",Time,Date
	insertConditionTable "000011","SH","134",Time,Date
	insertConditionTable "000012","SH","134",Time,Date
	closeConnection()		'''�ر����ݿ�
End Sub

'''���ڲ���ƽ��
Sub TestClose()
	closeYestodayOrder()
	clearOpenedTable()
	clearTmpconditionTable()
End Sub


Sub TestSucceedClose()
	openConnection()
	count = HistoryArray.Count
	for i = 0 to count-1
		hp = HistoryPrice.GetAt(i)
		hc = HistoryCondition.GetAt(i)
		profit  = 0
		insertClosedTable Code,"SZ",hp,0,"0",profit,Time,Date,hc
	next
	closeConnection()
End Sub


'''����ΪVBA������ÿ1��������һ��
Sub APPLICATION_VBAStart()
    Call application.Settimer(0,1000)
End Sub



Sub APPLICATION_Timer(ID)

	if cdate(time)="09:15:01" then	    '''�������ڻ������
		envirnmentCheck()
	elseif cdate(time)="09:30:01" then  '''�����ݿ⣬�ر����쿪�ֵ����й�Ʊ������������������
		
		closeYestodayOrder()			'''�ر����쿪�Ķ൥
		clearOpenedTable()              '''�൥�رպ󽫿������������
	elseif cdate(time)=cdate("10:45:01") then  '''��9.45�ֵ�ʱ��ɸѡ��ȫ������1�Ĺ�Ʊ
		count = runK1()
		if count = 0 then
			logString("[ ��ʾ: ] ��������û�п�ѡ�Ĺ�Ʊ���н���")
		else 
			logString("[ ��ʾ: ] ɸѡ���˺�����[1]�Ĺ�Ʊ " & count & "ֻ")
		end if
    elseif cdate(time)=cdate("10:47:01") then  '''��10.00�ֵ�ʱ��ɸѡ��ȫ������12�Ĺ�Ʊ
    	count = runK12()
    	if count = 0 then
			logString("[ ��ʾ: ] û�з�������[12]�Ĺ�Ʊ")
		else 
			logString("[ ��ʾ: ] ɸѡ���˺�����[12]�Ĺ�Ʊ " & count & "ֻ")
		end if
    elseif cdate(time)=cdate("10:48:01") then  '''��10.15�ֵ�ʱ��ɸѡ��ȫ������13,123�Ĺ�Ʊ,�������û����Ҫ����Ĺ�Ʊ������������м�����
        count = runK123()
        if count = 0 then
			logString("[ ��ʾ: ] û�з�������[123]�Ĺ�Ʊ")
		else 
			logString("[ ��ʾ: ] ɸѡ���˺�����[123]�Ĺ�Ʊ " & count & "ֻ")
		end if
    	firstBuy()						'''�ȼ��������123�Ľ��й����������13����
    	count = runK13()
    	if count = 0 then
			logString("[ ��ʾ: ] û�з�������[13]�Ĺ�Ʊ")
		else 
			logString("[ ��ʾ: ] ɸѡ���˺�����[13]�Ĺ�Ʊ " & count & "ֻ")
		end if
    elseif cdate(time)=cdate("10:49:01") then  '''��10.30�ֵ�ʱ��ѡ������124��134�Ĺ�Ʊ�󣬲������û����Ҫ����Ĺ�Ʊ������������м�����
    	count = runK124()
    	if count = 0 then
			logString("[ ��ʾ: ] û�з�������[124]�Ĺ�Ʊ")
		else 
			logString("[ ��ʾ: ] ɸѡ���˺�����[124]�Ĺ�Ʊ " & count & "ֻ")
		end if
    	count = runK134()
    	if count = 0 then
			logString("[ ��ʾ: ] û�з�������[134]�Ĺ�Ʊ")
		else
			logString("[ ��ʾ: ] ɸѡ���˺�����[134]�Ĺ�Ʊ " & count & "ֻ")
		end if
    	secondBuy()
    elseif cdate(time)=cdate("10:50:00") then  '''�������Ƿ��Ѿ�������ɣ����������ɣ��ͷŸ��ͷŵĶ��������ҹر����ݿ�
    	logString("[ ��ʾ: ] �������ɵ���ʱ����")
    	removeStage1()
    	removeStage2()
		removeStage3()
    end if
End Sub

Sub APPLICATION_VBAEnd()
	call application.killtimer(0)
End Sub


Function getHistoryPrice(Code)
	Dim p
	p = 0
	count = HistoryArray.Count
	for i = 0 to count-1
		if Code = HistoryArray.GetAt(i) then
			p = HistoryPrice.GetAt(i)
			exit for
		end if
	next
	getHistoryPrice = p
End Function


Function getCondtion2(Code)
	Dim p
	p = 0
	count = HistoryArray.Count
	for i = 0 to count-1
		if Code = HistoryArray.GetAt(i) then
			p = HistoryCondition.GetAt(i)
			exit for
		end if
	next
	getCondtion2 = p
End Function

'''�����¼����������൥����ɹ��󴥷����¼�
Sub ORDER_OrderStatusEx(OrderID, Status, Filled, Remaining, Price, Code, Market, OrderType, Aspect, Kaiping)

	'''debugString(Status & ":" & Filled & ":"  & Remaining & ":"  & Price & ":"  & Code & ":"  & Market & ":" & OrderType & ":"  & Aspect & ":"  & KaiPing)
	if(Aspect = 0 and Kaiping=0 and Status="Filled") then '''����൥�Ѿ�ȫ���ɽ����������ݿ�
		openConnection()
		insertOpenedTable Code,Market,Price,Filled,Time,Date
		closeConnection()
	elseif (Aspect = 1 and Kaiping=0 and Status="Filled") then '''�����Ѿ�ƽ�˶൥���������ݿ�
			hp = getHistoryPrice(Code)
			hc = getCondtion2(Code)
			if hp>0 then
				openConnection()
				insertClosedTable Code,Market,hp,Price,Filled,Price-hp,Time,Date,hc
				closeConnection()
			end if
	end if
End Sub

