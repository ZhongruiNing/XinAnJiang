 !++++++++++++++三水源滞后演算模型中的河网汇流模块++++++++++++++++++++++++++++
!-----------输入：参数CS1,L1; 系列长度Ntime;单元流域个数nsub
!-----------输入：地表流量qsp1,壤中流流量qip1,地下径流流量qgp1
!-----------输入：初始时刻河网出流量Qin1
!-----------输出:  河网出流量Qinp1
!----------中间变量：Qtot
!--------- Program by Yingchun Huang on 2020-12-20-------------------        
    
  Subroutine Lagnetworknew(CS1,L1,ntime,nsub,qsp1,qip1,qgp1,Qin1,Qinp1)
        Implicit none
        Integer ntime,nsub,itime,isub
        Real CS1(nsub)
        Integer L1(nsub)
        Real qsp1(ntime,nsub),qip1(ntime,nsub),qgp1(ntime,nsub)
        Real Qin1(nsub),Qinp1(ntime,nsub)
        
        Real Qtot(ntime,nsub) 
        Integer i
        print *,ntime,nsub
        
        
        do isub=1,nsub
        
     !=======初始时刻径流量赋值======    
            Qinp1(1,isub)=Qin1(isub)
            do i=1,L1(isub)         !滞后时段的流量赋值
                Qinp1(i,isub)=Qin1(isub)
            enddo
        
            do itime=1,ntime
                Qtot(itime,isub)=qsp1(itime,isub)+qip1(itime,isub)+qgp1(itime,isub)
            enddo
    !===============河网汇流计算==========        
            do itime=2+L1(isub),ntime
       
                Qinp1(itime,isub)=CS1(isub)*Qinp1(itime-1,isub)+(1.0-CS1(isub))*Qtot(itime-L1(isub),isub)                                                                                                
            enddo
        
        
        enddo
        
        
        Return
  
  End  Subroutine LagNetworknew