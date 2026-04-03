!++++++++++++++马斯京根河道流量演算模型++++++++++++++++++++++++++++
!-----------输入：参数MX1,MN1------------
!-------------演算时段长DeltaT;系列长度Ntime,单元流域个数nsub--------------
!------------------上游流量过程Qinp1; 下断面起始流量Qout1    
!-----------输出：下游流量过程Qoutp1----------------
!----------中间变量： C0, C1, C2--------------
!--------- Program by Yingchun Huang 2020-12-20-------------------
    
    Subroutine Muskingumnew(MX1,MN1,DeltaT,ntime,nsub,Qinp1,Qout1,Qoutp1,Qoutlet)   
            Implicit none
            Integer ntime,nsub,itime,isub
            Real MX1(nsub)
            Integer MN1(nsub)
            Real Qout1(nsub)
            Real Qoutp1(ntime,nsub)
            Real Qinp1(ntime,nsub)
            Real Qoutlet(ntime)
            
            Real DeltaT  
            Real Mk,Mx,n            
            Real C0,C1,C2
            Integer j 
 
            
            do isub=1,nsub
                Mk=DeltaT
                Mx=MX1(isub)
                n=MN1(isub)
!=========计算中间变量C0, C1, C2================      
                C0=(0.5*DeltaT-Mk*Mx)/(0.5*DeltaT+Mk-Mk*Mx)
                C1=(0.5*DeltaT+Mk*Mx)/(0.5*DeltaT+Mk-Mk*Mx)
                C2=1.0-C0-C1
 !========逐河段马斯京根演算===================
                Qoutp1(1,isub)=Qout1(isub)    ! 下游河段起始时刻流量赋值
                do j=1,n                                    ! 逐河段       
                    do itime=2,ntime                  ! 从第二个时刻开始逐时段演算
                        Qoutp1(itime,isub)=C0*Qinp1(itime,isub)+C1*Qinp1(itime-1,isub)+C2*Qoutp1(itime-1,isub)    !下游河段出流过程
                   enddo
                    If (j .lt. n) then      !判断若不是最后一个河段（j比n小），上一个河段出流是下一个河段的入流
                        do itime=2,Ntime
                            Qinp1(itime,isub)= Qoutp1(itime,isub)  
                        enddo
                    endif
                enddo

            
                    if (n .le. 0) then              !若河段数为0，不进行演算
                        do itime=2,ntime
                            Qoutp1(itime,isub)=Qinp1(itime,isub)
                        enddo
                    endif
        enddo     
            
!----------------------输出流域出口断面流量过程-     
            Qoutlet=0.0
            
                   do itime=1,ntime
                       do isub=1,nsub
                           Qoutlet(itime)=Qoutlet(itime)+Qoutp1(itime,isub)
                       enddo
                   enddo
                   
    Return        
    End Subroutine Muskingumnew
    