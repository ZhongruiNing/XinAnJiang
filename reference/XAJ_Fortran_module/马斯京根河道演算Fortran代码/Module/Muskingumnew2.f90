!++++++++++++++马斯京根河道流量演算模型++++++++++++++++++++++++++++
!--------------输入1（模型参数文件）：马斯京根法参数MN1、MX1、MK1，演算时段长DeltaT、系列长度ntime、单元流域个数nsub 
!--------------输入2（等时段水位流量输入文件）： 各河段对应区间的时序流量输入Qinp1
!--------------输入3（模型起始状态文件）： 各河段下断面的初始流量Qoutini   
!-----------输出1：所有节点的流量过程Qtot----------------
    
    
!----------中间变量： C0, C1, C2--------------
!--------- Program by Yingchun Huang 2021-09-04-------------------
    
    Subroutine Muskingumnew2(nsub,ntime,DT,REACD,SECTN,REASCD,REAECD,RVMSKN,RVMSKXE,RVMSKKE,Qinp1,Qoutini,Qtot)
     
            Implicit none
            Integer ntime,itime
            Integer nsub,isub
            Real DeltaT,DT
            Character REACD(nsub)*100
            Integer SECTN(nsub),REASCD(nsub),REAECD(nsub),RVMSKN(nsub)
            Real RVMSKXE(nsub),RVMSKKE(nsub)
            
            Real Qinp1(ntime,nsub)
            Real Qoutini(nsub)
          
            Real Qin(ntime),Qout(ntime),Qout1
            
            Real Qsum(ntime,nsub+1)
            Real Qtot(ntime,nsub+1)
            Real Mk,Mx 
            Integer n

            
 
            DeltaT=DT/3600.0   !----将时间单位由秒转化为小时 
            Qsum=0.0    !流量先假定为0
            Qtot=0.0      !流量先假定为0
            
            
            !-----首先将区间产流量加在河流的下断面
            do isub=1,nsub
                do itime=1,ntime
                    Qsum(itime,REAECD(isub))=Qsum(itime,REAECD(isub))+Qinp1(itime,(isub))
                enddo
            enddo
         
 
            do isub=1,nsub       !所有单元流域循环进行河道汇流演算
                   
                    !--------参数赋值
                    Mk=RVMSKKE(REASCD(isub))
                    Mx=RVMSKXE(REASCD(isub))
                    n=RVMSKN(REASCD(isub))
                    Qout1=Qoutini(REASCD(isub))
                    
                    do itime=1,ntime
                        Qin(itime)=Qsum(itime,REASCD(isub))     !---上断面总流量=上游来水演算结果 

                    enddo
                    
                    call Muskingum(Mk,Mx,n,DeltaT,Qin,Qout1,ntime,Qout)   !---调用马斯京根方法进行演算
     
                    
                    do itime=1,ntime
                        Qsum(itime,REAECD(isub))=Qsum(itime,REAECD(isub))+Qout(itime)     !--下段面总流量=区间产流量+上游来水在下断面的演算结果累加
                       
                        Qtot(itime,REASCD(isub))=Qin(itime)      !该子流域所对应节点（上断面）的总流量
                       
                    enddo
            enddo
            
            !---------整个流域出口位置节点(nsub+1)的流量过程
            do itime=1,ntime
                Qtot(itime,nsub+1)=Qsum(itime,nsub+1)
            enddo
                
    Return        
    End Subroutine Muskingumnew2
    