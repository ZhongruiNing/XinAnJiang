Attribute VB_Name = "foreshow"
Sub foreshows()
ReDim glpwhf(LongTime), glqcalhf(LongTime), glqobshf(LongTime), glqadjhf(LongTime)
  
  Call findtime(CountFlood)
  NoFlood = FloodNo(1, CountFlood)
  Call daytime
 
  Call readfore

  yb09.Show


End Sub
