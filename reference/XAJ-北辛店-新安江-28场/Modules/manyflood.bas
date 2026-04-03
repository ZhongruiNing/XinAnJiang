Attribute VB_Name = "manyflood"
Sub manyfloods()

  For CountFlood = 1 To NumberNo
     Call findtime(CountFlood)
     NoFlood = FloodNo(1, CountFlood)
     Call zmqbbb(CountFlood)
     Call daytime
     'Call fenpjunj
     'Call ybchxxx
 Next CountFlood

End Sub


