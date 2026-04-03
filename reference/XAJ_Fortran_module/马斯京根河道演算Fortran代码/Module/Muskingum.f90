!+++++++++++马斯京根河道流量演算模型++++++++++++++++++++++++++++
!-----------输入：参数Mk,Mx,n 
!-----------演算时段长DeltaT;上游流量过程Qin; 下断面起始流量Qout1,系列长度ntime--------------
!-----------输出：下游流量过程Qout----------------
!-----------中间变量： C0, C1, C2--------------
!-----------Program by Yingchun Huang 2021-09-04-------------------
    
Subroutine Muskingum(Mk,Mx,n,DeltaT,Qin,Qout1,ntime,Qout)
    Implicit none
    Real Mk,Mx,DeltaT  
    Integer n
    Integer ntime,itime
    Real Qout1
    Real Qin(Ntime)
    Real Qout(Ntime)
            
!----------中间变量--------------                 
    Real C0,C1,C2
    Integer j 
            
!=========计算中间变量C0, C1, C2================      
    C0=(0.5*DeltaT-Mk*Mx)/(0.5*DeltaT+Mk-Mk*Mx)      !该行代码和原来一致未变
    C1=(0.5*DeltaT+Mk*Mx)/(0.5*DeltaT+Mk-Mk*Mx)    !该行代码和原来一致未变
    C2=1.0-C0-C1     
            
!========逐河段马斯京根演算===================
    Qout=0.0         !所有时刻流量先赋0
    Qout(1)=Qout1         ! 下游河段起始时刻流量赋值
    do j=1,n                     !逐河段                
        do itime=2,ntime         !从第二个时刻开始逐时段演算
            Qout(itime)=C0*Qin(itime)+C1*Qin(itime-1)+C2*Qout(itime-1)      !下游河段出流过程
        enddo               
        If (j .lt. n) then      !判断若不是最后一个河段（j比n小），上一个河段出流是下一个河段的入流
            do itime=2,ntime
                Qin(itime)=Qout(itime)          
            enddo
        endif
    enddo

    if (n .le. 0) then              !若河段数为0，不进行演算
        do itime=2,ntime
            Qout(itime) =Qin(itime)
        enddo
    endif

    Return
End Subroutine Muskingum  
      