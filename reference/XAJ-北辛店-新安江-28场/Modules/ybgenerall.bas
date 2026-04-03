Attribute VB_Name = "ybgenerall"


Sub ybgenera()
   Dim Na As Integer, naa As Integer
   
   Call findnumber
 
  For CountFlood = 1 To NumberNo
     Call findtime(CountFlood)
     NoFlood = FloodNo(1, CountFlood)
     Call zmqbbb(CountFlood)
     Call findtime(CountFlood)
     Call timeint
     Call timedayz
     ' Call calcu_pan
     'Call calcu_pan
     'Call calcu_averagp_period
     Call chyubasH
 Next CountFlood
 
End Sub


