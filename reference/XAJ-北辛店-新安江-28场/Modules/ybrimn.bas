Attribute VB_Name = "ybrimn"
Sub rmxjss()

   Call findnumberd
   

  For CountFlood = 1 To NumberNo
     Call findtimed(CountFlood)
     NoFlood = FloodNo(1, CountFlood)
     Call daytimed(LongTimeD)
     Call chyubasD
      
  Next CountFlood
End Sub

Sub rmxwater()

   Call findnumberd
   dyly = "lm"

  For CountFlood = 1 To NumberNo
     Call findtimed(CountFlood)
     NoFlood = FloodNo(1, CountFlood)
     Call daytimed(LongTimeD)
     Call WatersurfaceD
    
  Next CountFlood
End Sub

