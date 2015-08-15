

'Goals:
'Downlaod stock quote data from yahoo finance, calculate day over day equity returns
'Use empirical distribution of equity returns to project paths for stock prices using monte carlo 
'Use monte carlo projections to  calculate option prices and option implied volatility surface ( by inverting black scholes equation)
'Optimal strategy for option trading
'Must be implemented in Excel platform hence excel VBA



Option Explicit

Public Save_Path As String, Start_Date As Date, End_Date As Date, M As Variant, D As Variant, Y As Integer, N_Min As Lon
Public URL_Root As String, Stock As String, URL As String, Stock_bis As String
Public Start_String As String, End_String As String, Data_Folder As String, DownloadState As Boolean
Public Weights() As Double



Sub Form_Porfolio()

'Sub downloads stock prices from yahoo finance: list of stock is read from excell spreadsheet

Dim iStock As Integer
Data_Folder = CStr(Range("Data_Folder").Value)

'checking that data folder exists

If Len(Dir(Data_Folder, vbDirectory)) = 0 Then
 
        MsgBox ("Data Folder Doesn't Exist")
    
        End
        
End If

URL_Root = "http://table.finance.yahoo.com/table.csv?"
Start_Date = Range("Start_Date").Value: End_Date = Range("End_Date").Value

For iStock = 1 To CInt(Range("Num_Stocks").Value)  ' loop over all stock ticker

     Stock = CStr(Range("Stock_List").Cells(1 + iStock, 1).Value): Stock_bis = Stock: Stock = "s=" + Stock

     Get_Start_String
     Get_End_String

     URL = URL_Root + Stock + Start_String + End_String + "&g=d"   ' forming URL for quote prices download

     Get_Data_From_Yahoo URL, Data_Folder & Stock_bis & ".csv"
     
     If DownloadState = False Then
     
             Range("DownloadState").Cells(iStock + 1, 1).Value = "Failed"
     Else
             Range("DownloadState").Cells(iStock + 1, 1).Value = "Success"
           
     End If
     
     
Next iStock

 MsgBox ("Finished Download of Stock Quotes")


End Sub


Sub Get_Start_String()


Start_String = ""

    M = Application.WorksheetFunction.Max(Month(Start_Date) - 1, 1): D = Application.WorksheetFunction.Max(Day(Start_Date), 1): Y = Year(Start_Date)

    Start_String = Start_String + "&a=" + CStr(M)
    Start_String = Start_String + "&b=" + CStr(D)
    Start_String = Start_String + "&c=" + CStr(Y)
       

End Sub

Sub Get_End_String()

End_String = ""

    M = Application.WorksheetFunction.Max(Month(End_Date) - 1, 1): D = Application.WorksheetFunction.Max(Day(End_Date), 1): Y = Year(End_Date)

    
    If (M = "" Or D = "") Then
    
          MsgBox ("Wrong Start or/and End Dates")
       End
   
   End If
   
    End_String = End_String + "&d=" + CStr(M)
    End_String = End_String + "&e=" + CStr(D)
    End_String = End_String + "&f=" + CStr(Y)
       

End Sub



Sub Get_Data_Points()


' Calculate the number of trading days for each stock file


Dim i As Integer, File_Path As String, N() As Variant
Dim Line As String

Data_Folder = CStr(Range("Data_Folder").Value)

ReDim N(CInt(Range("Num_Stocks").Value))

For i = 1 To CInt(Range("Num_Stocks").Value)

 
File_Path = Data_Folder + CStr(Range("Stock_List").Cells(i + 1, 1).Value) + ".csv"
  
   
     N(i - 1) = 0
  
            Open File_Path For Input As #1
  
                  Do While Not (EOF(1))
  
                  Line Input #1, Line
  
                  N(i - 1) = N(i - 1) + 1
  
  
           Loop

       Range("Num_Data_Points").Cells(i + 1, 1).Value = N(i - 1)
       Close #1

Next i


N_Min = CLng(Application.WorksheetFunction.Min(N)) - 1 ' this is minimum # of trading days (-1) To take care of header



End Sub


Sub Get_Weights()

  Dim i As Integer

  ReDim Weights(CInt(Range("Num_Stocks").Value)) As Double
  
  
  For i = 1 To CInt(Range("Num_Stocks").Value)

   Weights(i - 1) = CDbl(Range("Weight").Cells(i + 1, 1).Value)
 
  Next i


End Sub



Sub Get_Portfolio()

  Dim i As Integer, j As Integer
  Dim xc() As String
  
  Dim Price As Double, data As String
  Dim Line As String, LineData() As String, File_Path As String, seperator As String, line2() As String

  
  Dim DataRange As Range
  
  
  Dim StockChart As Chart
  Dim Portfolio_Value As Variant, Series As Variant
  
 
  Range("Stat").ClearContents
  
  Get_Weights
  Get_Data_Points
  
  
  Set DataRange = Range(Range("ReturnData").Cells(1, 1), Range("ReturnData").Cells(N_Min, 2))
  Portfolio_Value = DataRange.Value
  Data_Folder = CStr(Range("Data_Folder").Value)
 
  
 
  For i = 1 To CInt(Range("Num_Stocks").Value)
      
      File_Path = Data_Folder + CStr(Range("Stock_List").Cells(i + 1, 1).Value) + ".csv"
      Open File_Path For Input As #i
      
      Line Input #i, Line   ' this is header line
  
  Next i
  
  
  Range("TimeRange").ClearContents
  
  
  For i = 1 To N_Min - 1
       
        
        Price = 0
         
         For j = 1 To CInt(Range("Num_Stocks").Value)
                    
                   Line Input #j, Line
                   
                   LineData = Split(Line, ",")   ' Date Open High Low Close Volume AdjustedClose
                   
                   data = CDbl(LineData(4))     ' this is closing price for stock
                   
                   Price = Price + Weights(j - 1) * CDbl(data)
                       
         Next j
               
             Portfolio_Value(i, 2) = CDbl(Price)
             Portfolio_Value(i, 1) = CDate(LineData(0))
    
    
  Next i
    
           DataRange.Value = Portfolio_Value
            
           Set Series = Worksheets("ControlPanel").ChartObjects("Chart 5").Chart.SeriesCollection("Portfolio Value")   '
                
                     Series.Values = Range(DataRange.Cells(1, 2), DataRange.Cells(N_Min, 2))
                     Series.XValues = Range(DataRange.Cells(1, 1), DataRange.Cells(N_Min, 1))
                     
                     
           Worksheets("ControlPanel").ChartObjects("Chart 5").Chart.Axes(xlCategory).MinimumScale = Application.WorksheetFunction.Min(Range(DataRange.Cells(1, 1), DataRange.Cells(N_Min, 1)))
    
           Worksheets("ControlPanel").ChartObjects("Chart 5").Chart.Axes(xlCategory).MaximumScale = Application.WorksheetFunction.Max(Range(DataRange.Cells(1, 1), DataRange.Cells(N_Min, 1)))
   
   
             
          
          For i = 1 To CInt(Range("Num_Stocks").Value)
          
             Close #i
          
          Next i
    
       
          Stat_Analysis Range(Range("ReturnData").Cells(1, 2), Range("ReturnData").Cells(N_Min, 2)).Value


End Sub



Sub Stat_Analysis(Input_Data)

   Dim Output_Data() As Variant
   Dim N As Integer, i As Integer, NumPercentiles As Long, Nrow As Long

   Dim PercentileValues As Variant

   N = UBound(Input_Data)
   NumPercentiles = Range("NumPercentiles").Value



   Range("Returns").ClearContents
   Output_Data = Range(Range("Return").Cells(2, 1), Range("Return").Cells(N - 1, 1)).Value

   
   For i = 1 To N - 2
 
        Output_Data(i, 1) = Input_Data(i, 1) / Input_Data(i + 1, 1) - 1      'geometric return
       
   Next i


    
    Range(Range("Return").Cells(2, 1), Range("Return").Cells(N + 1, 1)).Value2 = Output_Data


    Range("Mean").Value = Application.WorksheetFunction.Average(Output_Data)
    Range("Median").Value = Application.WorksheetFunction.Median(Output_Data)
    Range("Volatility").Value = Application.WorksheetFunction.StDev(Output_Data)
   
    Range("Skew").Value = Application.WorksheetFunction.Skew(Output_Data)
    Range("Kurtosis").Value = Application.WorksheetFunction.Kurt(Output_Data)

    PercentileValues = Range(Range("Percentiles").Cells(2, 1), Range("Percentiles").Cells(NumPercentiles + 1, 2)).Value
    Range("PercentileRange").ClearContents

   
   For i = 0 To NumPercentiles - 1
   
            PercentileValues(i + 1, 1) = i / (NumPercentiles - 1)
            
            PercentileValues(i + 1, 2) = Application.WorksheetFunction.Percentile(Output_Data, i / (NumPercentiles - 1))
   
   Next i
           
         
           Range(Range("Percentiles").Cells(2, 1), Range("Percentiles").Cells(NumPercentiles + 1, 2)).Value = PercentileValues
             
    With Worksheets("ControlPanel").ChartObjects("Chart 3").Chart
            
          .SetSourceData Range(Range("Percentiles").Cells(2, 1), Range("Percentiles").Cells(NumPercentiles + 1, 2))

    End With

End Sub



Sub Get_Data_From_Yahoo(URL As String, SaveFileAs As String)

'sub gets stock data from yahoo finance and save data in csv file

Dim WinHttpReq As Object
Dim Ostream As Object

Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")

WinHttpReq.Open "GET", URL, False
WinHttpReq.Send


If WinHttpReq.Status = 200 Then

        Set Ostream = CreateObject("ADODB.Stream")
        
        Ostream.Open
        Ostream.Type = 1
        Ostream.Write WinHttpReq.ResponseBody
    
             If Len(Dir(SaveFileAs)) <> 0 Then
        
                Kill (SaveFileAs)
    
             End If
    
        Ostream.SaveToFile (SaveFileAs)
        Ostream.Close
        DownloadState = True
End If



If WinHttpReq.Status = 404 Then
     
     MsgBox ("The Following Stock:  " & Stock_bis & "   was not found")
     DownloadState = False
     Exit Sub

End If
     

   
   AdjustStockFile SaveFileAs
 
   Set WinHttpReq = Nothing
   
   Set Ostream = Nothing
 
End Sub


Sub AdjustStockFile(FilePath As String)

    Dim InputDataMatrix() As String, iRow As Long
    Dim LineData As String, SplitLine() As String
  
  
    Open FilePath For Input As #1
         Line Input #1, LineData
         SplitLine = Split(LineData, vbLf)

   Close #1
  
         Kill (FilePath)
    
   
   Open FilePath For Output As #1
   
   
   For iRow = LBound(SplitLine) To UBound(SplitLine)
   
         Print #1, SplitLine(iRow);
         Print #1, ","
         
   Next iRow

         Close #1

End Sub

Function SelectStocks(Number_Of_Stocks As Long)

End Function


Sub SimulateReturns()

'here we use empirical distribution of stock return to project stock prices by monte carlo method 

Dim PercentileValues As Variant, NumPercentiles As Long, CPF As Double, Returns As Double, Price_0 As Double, Price As Double
Dim Seed As Double, iScenario As Long, NumScenarios As Long, NumSteps As Long, iStep As Long
Dim EquityReturns As Variant

NumPercentiles = CLng(Range("NumPercentiles").Value)

PercentileValues = Range(Range("Percentiles").Cells(2, 1), Range("Percentiles").Cells(NumPercentiles + 1, 2)).Value

'change seed for random number generator
 
Rnd (-10) : Randomize (CDbl(Range("Seed").Value))
    
Range("SiReturns").ClearContents
Range("EquityProjection").ClearContents
    
NumScenarios = Range("NumScenarios").Value
NumSteps = Range("NumSteps").Value

Price_0 = 1

EquityReturns = Range(Range("Scenarios").Cells(1, 1), Range("Scenarios").Cells(NumScenarios, NumSteps)).Value

For iScenario = 1 To Range("NumScenarios").Value
    
             Price = Price_0
          
             Application.StatusBar = "Running Scenario #   " & CStr(iScenario)
            
             DoEvents
          
                For iStep = 1 To NumSteps
             
                      CPF = Rnd  'cumulative probability function value: uniformly distributed
                      
                      Returns = Get_Return_From_CPF(CPF, PercentileValues) ' inverting CPF function, daily return
                      
                      Price = Price * (1 + Returns)  'stock prices, geometric projection
                      
                      EquityReturns(iScenario, iStep) = Price
                      
                      
                      ' If iStep = 1 Then
                       
                       '     Range("SimReturns").Cells(iScenario + 1, 1).Value = Returns
                            
                       'End If
                       
              
                Next iStep
          
        
          
        '  Range("SimReturns").Cells(iScenario + 1, 1).Value = Returns
          
Next iScenario
    
    
    Range(Range("Scenarios").Cells(1, 2), Range("Scenarios").Cells(NumScenarios, NumSteps + 1)).Value = EquityReturns
    
           

End Sub



Function Get_Return_From_CPF(CPF As Double, InputMatrix As Variant)


   'generate equity return scenario based on cumulative probability function value


   Dim iRow As Long, MaSize As Long, Alpha As Double
   
   MaSize = CLng(Range("NumPercentiles").Value)
    
    
    
    If CPF <= InputMatrix(1, 1) Then
    
            Get_Return_From_CPF = InputMatrix(1, 2)
            Exit Function
     End If
   
   
    If CPF >= InputMatrix(MaSize - 1, 1) Then
    
            Get_Return_From_CPF = InputMatrix(MaSize - 1, 2)
            Exit Function
     End If
   
   
   For iRow = 1 To MaSize - 1
   
             
          If CPF > InputMatrix(iRow, 1) And CPF < InputMatrix(iRow + 1, 1) Then
          
               
                   Alpha = (CPF - InputMatrix(iRow, 1)) / (InputMatrix(iRow + 1, 1) - InputMatrix(iRow, 1))
                    
                   Get_Return_From_CPF = InputMatrix(iRow, 2) + Alpha * (InputMatrix(iRow + 1, 2) - InputMatrix(iRow, 2))
          
                    Exit For
          End If
          
        
          If CPF = InputMatrix(iRow + 1, 1) Then
          
                   Get_Return_From_CPF = InputMatrix(iRow + 1, 2)
                   Exit For
          End If
          
   Next iRow
          
               
End Function



Sub Calculate_Option_Payoff()

'sub calculates option payoff for multiple strikes and expiration dates using monte carlo simulation of stock prices:  prices were stored in excell spreadsheet

Dim K_Min  As Double, K_Max As Double, K As Double, LocalPayOff As Double, IR As Double, SpotPrice As Double, OptionState As Double, IV As Double
Dim PayOff As Double
Dim EquityPrice As Variant, OptionData As Variant, IVData As Variant
Dim NumScenarios As Long, NumSteps As Long, iStrike As Long, NumStrikes As Long, iStep As Long, iScenario As Long
Dim OptionType As String
Dim OptionIV As Variant

 NumScenarios = Range("NumScenarios").Value
 NumSteps = Range("NumSteps").Value
 NumStrikes = Range("NumStrikeSteps").Value


' define range of strike prices

   K_Min = Range("K_Min").Value
   K_Max = Range("K_Max").Value
   
   IR = CDbl(Range("IR").Value) 'simply compounded short term interest rate
   
   SpotPrice = Range("SpotPrice").Value
   OptionType = Range("OptionType").Value
   Range("OptionData").ClearContents
   Range("IVData").ClearContents

   OptionData = Range(Range("OptionPrices").Cells(1, 1), Range("OptionPrices").Cells(NumStrikes, NumSteps)).Value
   IVData = Range(Range("OptionIVs").Cells(1, 1), Range("OptionIVs").Cells(NumStrikes, NumSteps)).Value

'load stock price projection ( > 1000 scenarios ) from excel spreadsheet

  EquityPrice = Range(Range("Scenarios").Cells(1, 2), Range("Scenarios").Cells(NumScenarios, NumSteps + 1)).Value

      For iStrike = 1 To NumStrikes ' loop on strik prices

                 K = K_Min + (iStrike - 1) * (K_Max - K_Min) / NumStrikes

                 Range("OptionPrices").Cells(iStrike, 1).Value = K
                 Range("OptionIVs").Cells(iStrike, 1).Value = K
        
                 Application.StatusBar = "Processing Strike number  " & CStr(iStrike)
          
                 DoEvents
            
            
                                  For iStep = 1 To NumSteps   ' loop on expiration date
                
                
                                                PayOff = 0


                                                        For iScenario = 1 To NumScenarios    ' calculate option payoff for all scenarios
                                     
                                                                LocalPayOff = SpotPrice * EquityPrice(iScenario, iStep) - K
                                             
                                             
                                                                    If OptionType = "Call" Then
                                                                           
                                                                           OptionState = 1
                                            
                                                                            PayOff = PayOff + Application.WorksheetFunction.Max(LocalPayOff, 0)
                                                                    Else
                                                                           OptionState = -1
                                                                           
                                                                            PayOff = PayOff + Application.WorksheetFunction.Max(-LocalPayOff, 0)
                                                                    End If
                                            
                                                        Next iScenario
                                     
                                     
                                                
                                                PayOff = PayOff / (NumScenarios * (1 + IR) ^ iStep)
                                                
                                                OptionData(iStrike, iStep) = PayOff
                                                
                                               OptionIV = Get_IV_From_Price(PayOff, SpotPrice, K, CDbl(IR), CDbl(iStep), OptionState)

                                            
                                            If OptionIV <> 0 Then
                                                
                                                    IVData(iStrike, iStep) = OptionIV
                                                
                                            Else
                                                  
                                                    IVData(iStrike, iStep) = ""
                                            
                                            End If
                                            
                                   
                                 Next iStep
            
      Next iStrike
 
          
         Range(Range("OptionPrices").Cells(1, 2), Range("OptionPrices").Cells(NumStrikes, NumSteps + 1)).Value = OptionData
         Range(Range("OptionIVs").Cells(1, 2), Range("OptionIVs").Cells(NumStrikes, NumSteps + 1)).Value = IVData

End Sub


Function Get_Option_Price(OptionType As Double, So As Double, K As Double, Sigma As Double, r As Double, T As Double) As Double

'sub uses black scholes equation for option pricing: call and put 


Dim d1 As Double, d2 As Double
 

d1 = (Log(So / K) + (r + 0.5 * Sigma ^ 2) * T) / (Sigma * Sqr(T))
d2 = (Log(So / K) + (r - 0.5 * Sigma ^ 2) * T) / (Sigma * Sqr(T))
   
Get_Option_Price = OptionType * (So * Application.WorksheetFunction.NormSDist(OptionType * d1) - K * (Exp(-r * T)) * Application.WorksheetFunction.NormSDist(OptionType * d2))
   
End Function


Function Get_IV_From_Price(Price_0 As Double, So As Double, K As Double, r As Double, T As Double, OptionType As Double) As Double

Dim SigmaMin As Double, SigmaMax As Double, Price As Double, Epsilon As Double, Tolerance As Double, Sigma As Double
    
    
If Price_0 <= CDbl(Range("OptionPriceThreshold").Value) Or Price_0 <= (So * Exp(r * T) - K) Then


         Get_IV_From_Price = 0
      
      Exit Function

End If

   
SigmaMin = 0: SigmaMax = 1000: Epsilon = 10: Tolerance = 0.0001
  

Do While Abs(Epsilon) > Tolerance
    
            Sigma = (SigmaMin + SigmaMax) / 2
         
    
             Epsilon = Price_0 - Get_Option_Price(OptionType, So, K, Sigma, r, T)
             
             If Epsilon > 0 Then
             
                     SigmaMin = Sigma
             End If
             
             If Epsilon < 0 Then
             
                      SigmaMax = Sigma
              End If
              
             If Epsilon = 0 Then
             
                    Exit Do
             End If
Loop

Get_IV_From_Price = Sigma * Sqr(252)   'annualized implied volatility
      
End Function


Function LocalHaltonSingleNumber(N As Double, b As Double) As Double


Dim n1 As Double, no As Double, hn As Double, f As Double, r As Double

no = N: hn = 0: f = 1 / b


Do While (no > 0)

       
        n1 = Application.WorksheetFunction.Floor(no / b, 1)
        
        r = no - n1 * b
        
        hn = hn + f * r
        
        f = f / b
        no = n1

Loop

LocalHaltonSingleNumber = hn
        
End Function

Sub ScreenStocks()

Dim FilePath   As String, Stock As String, CieName As String, LineData As String, SplitLineData() As String, Sticker As String
Dim iFile As Long
Dim Firstletter_Selection As Boolean


FilePath = CStr(Range("SP500Path").Value)

Sticker = Range("Sticker").Value  ' this is first letter of stock sticker

Firstletter_Selection = CBool(Range("Firstletter_Selection").Value)

Open FilePath For Input As #2
Line Input #2, LineData   ' this is header

Do While Not (EOF(2))
   
   
            Line Input #2, LineData
            
            SplitLineData = Split(LineData, ",")
            
            Stock = SplitLineData(2): CieName = SplitLineData(1)
            
        
        If Firstletter_Selection = True Then
        
                  If (Left(Stock, 1)) = Sticker Then
            
                          Range("Stock_List").Cells(2, 1).Value = Stock
            
                          Range("Num_Stocks").Value = 1
            
                          Worksheets("ControlPanel").ChartObjects("Chart 5").Chart.ChartTitle.Text = CieName
                           DoEvents
            
            
                          Call Form_Portfolio
            
            
                               If DownloadState = True Then
            
                                    Call Get_Portfolio
                                    Application.Calculate
                                    
                                End If
                 
                          DoEvents
    
                  End If
            
            
            
            
        Else
            
                
            
                     Range("Stock_List").Cells(2, 1).Value = Stock
            
                     Range("Num_Stocks").Value = 1
            
                     Worksheets("ControlPanel").ChartObjects("Chart 5").Chart.ChartTitle.Text = CieName
                     
                     DoEvents
            
            
                     Call Form_Portfolio
            
            
                  If DownloadState = True Then
            
                     Call Get_Portfolio
                     
                     Application.Calculate
                 
                
                     'MsgBox (" Next stock? ")
            
                  End If
                 
                    DoEvents
    
       End If
    
    Loop
            
            
   Close #2


End Sub



Sub ConvergenceTest()

    
Dim MaxScenarios As Long, i As Long, res1 As Double, res2 As Double, res3 As Double, res4 As Double, res5 As Double
Dim MasterData() As Variant, OutputData() As Variant
Dim NumScenarios  As Long, NumSteps As Long, SkipStep As Long


Range("ConvergenceData").ClearContents
SkipStep = Range("SkipStep").Value
NumScenarios = Range("NumScenarios").Value
NumSteps = Int(NumScenarios / SkipStep)


ReDim OutputData(NumSteps, 5)


For i = SkipStep To Range("NumScenarios").Value Step SkipStep


             MasterData = Range(Range("Scenarios").Cells(1, 2), Range("Scenarios").Cells(i, 2)).Value
    
    
             res1 = Application.WorksheetFunction.Average(MasterData) - 1
             res2 = Application.WorksheetFunction.Median(MasterData) - 1
             res3 = Application.WorksheetFunction.StDev(MasterData)
             res4 = Application.WorksheetFunction.Kurt(MasterData)
             res5 = Application.WorksheetFunction.Skew(MasterData)
             
             OutputData(i / SkipStep - 1, 0) = i
             OutputData(i / SkipStep - 1, 1) = res1
             OutputData(i / SkipStep - 1, 2) = res2
             OutputData(i / SkipStep - 1, 3) = res3
             OutputData(i / SkipStep - 1, 4) = res4
             OutputData(i / SkipStep - 1, 5) = res5
            
 
             Application.StatusBar = "Processing scenarios " & CStr(i)
             DoEvents
     
 Next i
        
             Range(Range("Convergence").Cells(1, 1), Range("Convergence").Cells(NumSteps, 6)).Value = OutputData
  


End Sub
