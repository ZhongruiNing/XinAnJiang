!****************************************************************************
    Program  MainMultisub
        implicit none
 
        Integer ntime,itime
        Integer nsub,isub
        Real DeltaT,DT
   
        Character,allocatable::REACD(:)
        Integer,allocatable::SECTN(:),REASCD(:),REAECD(:),RVMSKN(:)
        Real,allocatable::RVMSKXE(:),RVMSKKE(:)
        
        Real,allocatable:: Qinp1(:,:),Qoutini(:),Qtot(:,:)
            
 !----------------------读取单元流域个数、模拟数据系列长度，模拟时段长DT（单位为秒）    
            open(11,file='configuration.txt')  
            read(11,*,end=101)     !读取题头行
            read(11,*,end=101)  nsub,ntime,DT
101     continue                
            close(11)
        
!----------------根据单元流域个数、数据长度，定义数组范围
            
          allocate( REACD(nsub),SECTN(nsub),REASCD(nsub),REAECD(nsub),RVMSKN(nsub),RVMSKXE(nsub),RVMSKKE(nsub))
          allocate(Qinp1(ntime,nsub),Qoutini(nsub),Qtot(ntime,nsub+1))
          
   
 !----------------------读取所有单元流域河段属性及参数信息
            open(12,file='muskingumpara.txt')
            read(12,*,end=102)  !读取题头行
            do isub=1,nsub
                read(12,*,end=102)  REACD(isub),SECTN(isub),REASCD(isub),REAECD(isub),RVMSKN(isub),RVMSKXE(isub),RVMSKKE(isub)
            print *,REASCD(isub),REAECD(isub)
            enddo
102         continue 
            close(12)
            
!----------------------读取所有单元流域入流过程
            open(13,file='muskinguminputdata.txt')
             read(13,*,end=103)   !读取题头行
             do itime=1,ntime
                 read(13,*,end=103)  (Qinp1(itime,isub),isub=1,nsub)
             enddo
103          continue
             close(13)
             
!----------------------读取所有单元流域初始时刻出流量         
             open(14,file='muskinguminitial.txt')
             read(14,*,end=104)     !读取题头行
             read(14,*,end=104) (Qoutini(isub),isub=1,nsub)
104          continue 
             close(14)
                 
   
               
  !-----------调用第六个微服务：马斯京根河道演算
            call   Muskingumnew2(nsub,ntime,DT,REACD,SECTN,REASCD,REAECD,RVMSKN,RVMSKXE,RVMSKKE,Qinp1,Qoutini,Qtot)
            
            
          
   !-----------输出第六个微服务计算结果1：河网汇流流量Qoutp1
            open(17,file='./output/XAJ_Muskingum.txt')             
            do itime=1,ntime
                write(17,2101)  (Qtot(itime,isub),isub=1,nsub+1)
            enddo
            close(17)                                       
            
2101        format(1x,100F14.2)
            
 STOP     
End Program  MainMultisub
    