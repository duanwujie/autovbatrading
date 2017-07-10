'该模块主要用来保存宏主执行主函数,不要拿做他用



public K1Array   '''存放了第1 根K线满足的股票
public K12Array  '''存放了第1，2根K线满足的股票
public K13Array  '''存放了第1，3根K先满足的股票
public K1Market
public K12Market
public K13Market


public K123Array '''第3根K线出来以后检查'
public K134Array '''第4根K线出来以后检查
public K124Array '''第4根K线出来以后检查

public K123Market
public K1342Market
public K1243Market





'''Set strNotice = "[ 提示: ]"
'''Set strWarning= "[ 警告: ]"
'''Set strError  = "[ 错误: ]"


public DBConnection		'''数据库连接'''

public SYSOpenRate		'''开仓费率'''


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


'''用于临时保留昨天购买的股票，价格和代码，用于计算利润和策略性能
Set HistoryArray = CreateObject("Stock.ArrayString")
Set HistoryPrice = CreateObject("Stock.Array")
Set HistoryCondition = CreateObject("Stock.ArrayString")



'''注意第一根K线是在9.45后形成的
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



'''用于选股时日志输出
Function debugString(String)
	call Document.DebugFile("C:\Weisoft Stock(x64)\Log\AutoTrader.TXT",String,1)
End Function

Function logString(String)
	call Document.DebugFile("C:\Weisoft Stock(x64)\Log\LogAutoTrader.TXT",String,1)
End Function

'''第1阶段，只是将123购买的删除
Function removeStage1()
	K123Array.RemoveAll()
	K123Market.RemoveAll()
End Function

'''第2阶段完成后，所有选出来的股票都应该买入了 
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

'''第3阶段完成后，所有买入的股票已经平仓了
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

'''该函数用于检查执行环境，看执行环境是否准许执行
Function envirnmentCheck()

	'''账户是否登录
	if Order.IsAccount("6004625") = 0 then
		debugString("注意要开盘了，但账户还没有登录，请登录")
	end if
	'''资金是否足够
	
	''''
	
	
	
	'''读取费率
	openConnection()
	SYSOpenRate = getOpenRate()
	closeConnection()	
	
	logString("[ 提示: ] 开仓费率:" & SYSOpenRate)
	
	 


End Function


'''K1 信号发生器
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


'''K1 and K2 信号发生器
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



'''K1 and K3 信号发生器
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



'''K1 and K2 and  K3 信号发生器
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

'''K1 and K2 and  K4 信号发生器
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


'''K1 and K3 and  K4 信号发生器
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



'''把所有满足K1条件的股票选举出来,第一个对时间要求不是很高，所以用了Application.PeekAndPump 
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



'''把所有满足K1 and K2条件的股票选举出来,第二个对时间要求不是很高，所以用了Application.PeekAndPump 
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

'''把所有满足K1 and K3条件的股票选举出来
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

'''把所有满足K1 and K2 and K3条件的股票选举出来
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


'''把所有满足K1 and K2 and K4条件的股票选举出来
Function runK124()
	Count = K12Array.Count
	for i=0 to Count-1
		Code = K12Array.GetAt(i)
		Market = K12Market.GetAt(i)
		Set Report1  = marketdata.GetReportData(Code,Market)
		if K124Signal(Code,Market)=1 and exclude123(Code)=1 then  '''这里排除123重复的
			K124Array.AddBack(Code)
			K124Market.AddBack(Market)
			openConnection()
			insertConditionTable Code,Market,"124",Time,Date
			closeConnection()
		end if
	next
	runK124 = K124Array.Count
End Function

'''把所有满足K1 and K3 and K4条件的股票选举出来
Function runK134()
	Count = K13Array.Count
	for i=0 to Count-1
		Code = K13Array.GetAt(i)
		Market = K13Market.GetAt(i)
		Set Report1  = marketdata.GetReportData(Code,Market)
		if K134Signal(Code,Market)=1 and exclude124(Code)=1 then   '''这里排除124重复的
			K134Array.AddBack(Code)
			K134Market.AddBack(Market)
			openConnection()
			insertConditionTable Code,Market,"134",Time,Date
			closeConnection()
		end if
	next
	runK134 = K134Array.Count
End Function



'''第一次购买的策略
Function firstBuy()

	Dim expendable_cash		'''当前可用资金
	Dim each_stock_cash		'''每只股票可用资金
	Dim count 				'''有几只股票
	
	Dim	stock_count			'''每只股票可买的股票数量
	
	
	count  = K123Array.Count
	if count>0 then			'''有股票可买才进行操作
		expendable_cash = Order.Account2(3,"")
		expendable_cash = expendable_cash / 2   '''这里刚开始只用所有资金的1半，后面可以自动配置
		
		expendable_cash = expendable_cash*(1-SYSOpenRate)   '''分成一半后，去掉手续费
		each_stock_cash = expendable_cash / count  '''计算每只股票可用资金
		
		for i=0 to count-1
			Code = K123Array.GetAt(i)
			Market = K123Market.GetAt(i)
			set Report1 = marketdata.GetReportData(Code,Market)
			stock_count  = each_stock_cash / (Report1.NewPrice)  '''这里要算上费率
			stock_count = Int(Int(stock_count)/100)*100    '''转换成标准手数
			
			rs = Order.Buy(1,stock_count,Report1.NewPrice,0,Code,Market,"",0)
			if rs <> -1 then
				logString("[ 提示: ]" & Code & " 以价格:" & Report1.NewPrice & " 买入：" & stock_count & "股")
			else
				logString("[ 错误: ] 购买 " & Code & " 失败")
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


'''第二次购买的策略
Function secondBuy()

	Dim expendable_cash		'''当前可用资金
	Dim each_stock_cash		'''每只股票可用资金
	Dim count1 				'''124有几只股票
	Dim count2              '''134有几只股票
	Dim count				'''124+134有几只股票
	Dim	stock_count			'''每只股票可买的股票数量
	
	count1  = K124Array.Count
	count2  = K134Array.Count
	count = count1+count2
	
	debugString(count & ":" & count1 & ":" & count2)
	
	if count>0 then			'''有股票可买才进行操作
		logString("[ 提示: ] 第4根K线有" & count & "只股票购买")
		expendable_cash = Order.Account2(3,"")
		
		expendable_cash = expendable_cash*(1-SYSOpenRate)   '''去掉手续费
		
		expendable_cash = expendable_cash  '''将所有资金用于购买股票
		each_stock_cash = expendable_cash / count  '''计算每只股票可用资金
		logString("each_stock_cash:" & each_stock_cash & " openrate:" & SYSOpenRate) 
	else
		logString("[ 提示: ] 第4根K线没有股票购买")
	end if
	
	
	
	
	'''购买满足124条件的股票
	for i=0 to count1-1
		Code = K124Array.GetAt(i)
		Market = K124Market.GetAt(i)
		set Report1 = marketdata.GetReportData(Code,Market)
		stock_count1  = each_stock_cash / Report1.NewPrice
		stock_count1 = Int(Int(stock_count1)/100)*100    '''转换成标准手数
		
		debugString("each_stock_cash:" & each_stock_cash  & "price:" & Report1.NewPrice & "stock_count1:" & stock_count1)

		rs = Order.Buy(1,stock_count1,Report1.NewPrice,0,Code,Market,"",0)
		if rs  <> -1 then
			logString("[ 提示: ]" & Code & " 以价格:" & Report1.NewPrice & " 买入：" & stock_count1 & "股")
		else
			logString("[ 错误: ] 购买 " & Code & " 失败")
		end if
	next
	
	'''购买满足134条件的股票
	for i=0 to count2-1
		Code = K134Array.GetAt(i)
		Market = K134Market.GetAt(i)
		set Report1 = marketdata.GetReportData(Code,Market)
		stock_count2  = each_stock_cash / Report1.NewPrice
		stock_count2 = Int(Int(stock_count2)/100)*100    '''转换成标准手数
		rs = Order.Buy(1,stock_count2,Report1.NewPrice,0,Code,Market,"",0)
		if rs <> -1 then
			logString("[ 提示: ]" & Code & " 以价格:" & Report1.NewPrice & " 买入：" & stock_count2 & "股")
		else
			logString("[ 错误: ] 购买 " & Code & " 失败")
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


'''这里将昨天购买的股票，以开盘后的价格平仓
Function closeYestodayOrder()
	removeStage3()			'''将stage3的状态初始化
	openConnection()		'''打开数据库
	strSQL = "SELECT * FROM opened"
	Set rs = CreateObject("ADODB.Recordset")
	rs.Open strSQL, DBConnection, , , adCmdText
	With rs
		if .EOF then
			debugString("[ 错误 ]:昨天的开盘表中没有数据，造成无法从数据库中读取代码来进行平仓")
			debugString("[ 警告 ]:尝试中交易账户中读取订单数据来进行平仓")
			emailSender "[ 错误 ]","昨天的开盘表中没有数据，造成无法从数据库中读取代码来进行平仓,请手动平仓"
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
	closeConnection()		'''关闭数据库
End Function




Function clearOpenedTable()
	openConnection()		'''打开数据库
	strSQL ="DELETE * from opened"
	DBConnection.Execute strSQL
	closeConnection()		'''关闭数据库
End Function


Function clearTmpconditionTable()
	openConnection()		'''打开数据库
	strSQL ="DELETE * from tmp_condition"
	DBConnection.Execute strSQL
	closeConnection()		'''关闭数据库
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

'''用于调试1条件
Sub K1()
	Count = runK1()
	debugString("已经筛选出条件为1的股票:" & Count & "只")
End Sub

'''用于调试12条件
Sub K2()
	Count = runK12()
	debugString("已经筛选出条件为12的股票:" & Count & "只")
End Sub


'''用于调试13, 123条件
Sub K3()
	Count = runK123()
	debugString("已经筛选出条件为123的股票:" & Count & "只")
	Count = runK13()
	debugString("已经筛选出条件为13的股票:" & Count & "只")
End Sub

'''用于调试124, 134条件
Sub K4()
	Count = runK124()
	debugString("已经筛选出条件为124的股票:" & Count & "只")
	Count = runK134()
	debugString("已经筛选出条件为134的股票:" & Count & "只")
End Sub



'''用于测试第1阶段的购买
Sub B1()
	firstBuy()
	removeStage1()
End Sub

'''用于测试第2阶段的购买
Sub B2()
	K124Array.AddBack("000425")
	K124Market.AddBack("SZ")
	
	openConnection()
	SYSOpenRate = getOpenRate()
	closeConnection()
	
	secondBuy()
	removeStage2()
End Sub

'''用于产生历史数据，进行测试,在opened,tmp_condition表中产生数据
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
	closeConnection()		'''关闭数据库
End Sub

'''用于测试平仓
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


'''设置为VBA启动后，每1秒钟运行一次
Sub APPLICATION_VBAStart()
    Call application.Settimer(0,1000)
End Sub



Sub APPLICATION_Timer(ID)

	if cdate(time)="09:15:01" then	    '''这里用于环境检查
		envirnmentCheck()
	elseif cdate(time)="09:30:01" then  '''打开数据库，关闭昨天开仓的所有股票，并且生成利润数据
		
		closeYestodayOrder()			'''关闭昨天开的多单
		clearOpenedTable()              '''多单关闭后将开单的数据清空
	elseif cdate(time)=cdate("10:45:01") then  '''在9.45分的时候筛选出全部满足1的股票
		count = runK1()
		if count = 0 then
			logString("[ 提示: ] 今天早盘没有可选的股票进行交易")
		else 
			logString("[ 提示: ] 筛选出了合条件[1]的股票 " & count & "只")
		end if
    elseif cdate(time)=cdate("10:47:01") then  '''在10.00分的时候筛选出全部满足12的股票
    	count = runK12()
    	if count = 0 then
			logString("[ 提示: ] 没有符合条件[12]的股票")
		else 
			logString("[ 提示: ] 筛选出了合条件[12]的股票 " & count & "只")
		end if
    elseif cdate(time)=cdate("10:48:01") then  '''在10.15分的时候筛选出全部满足13,123的股票,并检查有没有需要购买的股票，如果有则以市价买入
        count = runK123()
        if count = 0 then
			logString("[ 提示: ] 没有符合条件[123]的股票")
		else 
			logString("[ 提示: ] 筛选出了合条件[123]的股票 " & count & "只")
		end if
    	firstBuy()						'''先计算出满足123的进行购买后再生成13条件
    	count = runK13()
    	if count = 0 then
			logString("[ 提示: ] 没有符合条件[13]的股票")
		else 
			logString("[ 提示: ] 筛选出了合条件[13]的股票 " & count & "只")
		end if
    elseif cdate(time)=cdate("10:49:01") then  '''在10.30分的时候选出满足124，134的股票后，并检查有没有需要购买的股票，如果有则以市价买入
    	count = runK124()
    	if count = 0 then
			logString("[ 提示: ] 没有符合条件[124]的股票")
		else 
			logString("[ 提示: ] 筛选出了合条件[124]的股票 " & count & "只")
		end if
    	count = runK134()
    	if count = 0 then
			logString("[ 提示: ] 没有符合条件[134]的股票")
		else
			logString("[ 提示: ] 筛选出了合条件[134]的股票 " & count & "只")
		end if
    	secondBuy()
    elseif cdate(time)=cdate("10:50:00") then  '''这里检查是否已经购买完成，如果购买完成，释放该释放的东西，并且关闭数据库
    	logString("[ 提示: ] 清理生成的临时数据")
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

'''订单事件监听，当多单买入成功后触发该事件
Sub ORDER_OrderStatusEx(OrderID, Status, Filled, Remaining, Price, Code, Market, OrderType, Aspect, Kaiping)

	'''debugString(Status & ":" & Filled & ":"  & Remaining & ":"  & Price & ":"  & Code & ":"  & Market & ":" & OrderType & ":"  & Aspect & ":"  & KaiPing)
	if(Aspect = 0 and Kaiping=0 and Status="Filled") then '''这里多单已经全部成交，更新数据库
		openConnection()
		insertOpenedTable Code,Market,Price,Filled,Time,Date
		closeConnection()
	elseif (Aspect = 1 and Kaiping=0 and Status="Filled") then '''这里已经平了多单，更新数据库
			hp = getHistoryPrice(Code)
			hc = getCondtion2(Code)
			if hp>0 then
				openConnection()
				insertClosedTable Code,Market,hp,Price,Filled,Price-hp,Time,Date,hc
				closeConnection()
			end if
	end if
End Sub

