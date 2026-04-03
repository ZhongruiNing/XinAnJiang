 
!****************************************************************************
    Program  MainMultisub
        implicit none
        Real,allocatable:: Pobs1(:,:),Eobs1(:,:)
        Real,allocatable::  W1(:),WU1(:),WL1(:),S1(:),FR1(:)
        Real,allocatable:: K1(:),B1(:),C1(:),WM1(:),WUM1(:),WLM1(:),IM1(:),SM1(:),EX1(:),KG1(:),KI1(:)
        Real,allocatable:: CG1(:),CI1(:),CS1(:) 
        Integer,allocatable:: L1(:) 
        Real,allocatable:: Area1(:),QI1(:),QG1(:)
        Integer ntime,itime 
        Real DeltaT
 
        Real,allocatable:: wp1(:,:),wup1(:,:),wlp1(:,:),epp1(:,:),rp1(:,:),pep1(:,:)
        Real,allocatable:: sp1(:,:),rsp1(:,:),rip1(:,:),rgp1(:,:),eqq1(:,:),frp1(:,:)
        Real,allocatable:: qsp1(:,:),qip1(:,:),qgp1(:,:)
        Integer nsub,isub
        
        Real,allocatable:: Qin1(:),Qinp1(:,:),Qout1(:),Qoutp1(:,:)
 
        Real,allocatable:: ETP1(:,:),Qoutlet(:)
            
 !----------------------读取单元流域个数、模拟数据系列长度，模拟时段长DeltaT    
            open(11,file='configuration.txt')  
            read(11,*,end=100) 
            read(11,*,end=100)  nsub,ntime,DeltaT
100     continue                
            close(11)
        
!----------------根据单元流域个数、数据长度，定义数组范围
            allocate(Pobs1(ntime,nsub),Eobs1(ntime,nsub))
            allocate(W1(nsub),WU1(nsub),WL1(nsub),S1(nsub),FR1(nsub))
            allocate(K1(nsub),B1(nsub),C1(nsub),WM1(nsub),WUM1(nsub),WLM1(nsub),IM1(nsub),SM1(nsub),EX1(nsub),KG1(nsub),KI1(nsub))
            allocate(CG1(nsub),CI1(nsub),L1(nsub),CS1(nsub)) 
            allocate(Area1(nsub),QI1(nsub),QG1(nsub))
 
            allocate(wp1(ntime,nsub),wup1(ntime,nsub),wlp1(ntime,nsub),epp1(ntime,nsub),pep1(ntime,nsub),rp1(ntime,nsub))
            allocate(sp1(ntime,nsub),rsp1(ntime,nsub),rip1(ntime,nsub),rgp1(ntime,nsub),eqq1(ntime,nsub),frp1(ntime,nsub))
            allocate(ETP1(ntime,nsub)) 
            allocate(qsp1(ntime,nsub),qip1(ntime,nsub),qgp1(ntime,nsub))
            allocate(Qin1(nsub),Qinp1(ntime,nsub),Qout1(nsub),Qoutp1(ntime,nsub))
            allocate(Qoutlet(ntime))
  !----------------------读取所有单元流域降雨数据
            open(12,file='Precipitation.txt')
            read(12,*,end=99)
            do itime=1,ntime
                read(12,*,end=99) (Pobs1(itime,isub),isub=1,nsub)
            enddo
99     continue
            close(12)
    
!----------------------读取所有单元流域蒸发数据 
            open(13,file='Evapotranspiration.txt')
            read(13,*,end=98)
            do itime=1,ntime
                read(13,*,end=98) (Eobs1(itime,isub),isub=1,nsub)
            enddo
98      continue
            close(13)    
            
 !----------------------读取所有单元流域面积
            open(13,file='Area.txt')
            read(13,*,end=981)      
                read(13,*,end=98) (Area1(isub),isub=1,nsub)
981      continue
            close(13)      
            
 !----------------------读取所有单元流域的参数           
            open(15,file='XAJrunoffpara.txt')
                read(15,*,end=96)
                read(15,*,end=96) (K1(isub),isub=1,nsub)
                read(15,*,end=96) (B1(isub),isub=1,nsub)
                read(15,*,end=96) (C1(isub),isub=1,nsub)
                read(15,*,end=96) (WM1(isub),isub=1,nsub)
                read(15,*,end=96) (WUM1(isub),isub=1,nsub)
                read(15,*,end=96) (WLM1(isub),isub=1,nsub)
                read(15,*,end=96) (IM1(isub),isub=1,nsub)
                read(15,*,end=96) (SM1(isub),isub=1,nsub)
                read(15,*,end=96) (EX1(isub),isub=1,nsub)
                read(15,*,end=96) (KG1(isub),isub=1,nsub)
                read(15,*,end=96) (KI1(isub),isub=1,nsub)
                read(15,*,end=96) (CG1(isub),isub=1,nsub)
                read(15,*,end=96) (CI1(isub),isub=1,nsub)
                read(15,*,end=96) (L1(isub),isub=1,nsub)
                read(15,*,end=96) (CS1(isub),isub=1,nsub)

96      continue    
            close(15)

!----------------------读取所有单元流域的初始土壤状态值
            open(16,file='InitialSoil.txt')
                read(16,*,end=95)
                read(16,*,end=95) (W1(isub),isub=1,nsub)
                read(16,*,end=95) (WU1(isub),isub=1,nsub)
                read(16,*,end=95) (WL1(isub),isub=1,nsub)
                read(16,*,end=95) (S1(isub),isub=1,nsub)
                read(16,*,end=95) (FR1(isub),isub=1,nsub)
95      continue
            close(16)           
            
!----------------------读取所有单元流域初始壤中流径流深
            open(13,file='Qinflowini.txt')
            read(13,*,end=982)      
                read(13,*,end=98) (QI1(isub),isub=1,nsub)
982      continue
            close(13)   
            
 !----------------------读取所有单元流域初始地下径流深
            open(13,file='Qgroundini.txt')
            read(13,*,end=983)      
                read(13,*,end=98) (QG1(isub),isub=1,nsub)
983      continue
            close(13)  
            
!----------------------读取所有单元流域初始河网流量
            open(13,file='Qnetflowini.txt')
            read(13,*,end=984)      
                read(13,*,end=98) (Qin1(isub),isub=1,nsub)
984      continue
            close(13)  
  
            
 
            
            
!=======调用第一个微服务：蒸发能力计算微服务。根据观测蒸发值Eobs1、参数K，计算得到单元流域蒸散发能力ETP=======
            
            call ETPcal(Eobs1,K1,ETP1,ntime,nsub)
    
!-----------输出第一个微服务计算结果：各个单元流域各个时刻的蒸发能力ETP1             
            open(17,file='./output/XAJ_PotentialEvaporation.txt')             
            do itime=1,Ntime
                write(17,2100)  (ETP1(itime,isub),isub=1,nsub)
            enddo
            close(17)      
            

!==============================================================================    
 
     
!=======调用第二个微服务：新安江模型三层蒸散发与产流计算微服务。根据降雨Pobs1、蒸散发能力ETP1、土壤初始含水量W1,WU1,WL1、参数B1、C1、WM1、WUM1、WLM1、IM1，计算得到单元流域净雨量pep1,实际蒸发量epp1、产流量rp1,及土壤含水量状态wp1,wup1,wlp1    
    call XAJ3EvaYie(Pobs1,ETP1,W1,WU1,WL1,B1,C1,WM1,WUM1,WLM1,IM1,ntime,nsub,pep1,epp1,rp1,wp1,wup1,wlp1)
                            
!-----------输出第二个微服务计算结果1：各个单元流域各个时刻的实际蒸发量epp1  
            open(17,file='./output/XAJ_ActualEvaporation.txt')             
            do itime=1,Ntime
                write(17,2100)  (epp1(itime,isub),isub=1,nsub)
            enddo
            close(17)     
            
!-----------输出第二个微服务计算结果2：各个单元流域各个时刻的降雨扣除蒸发后的量pep1
            open(17,file='./output/XAJ_PreminusEta.txt')             
            do itime=1,Ntime
                write(17,2100)  (pep1(itime,isub),isub=1,nsub)
            enddo
            close(17)   
            
!-----------输出第二个微服务计算结果3：各个单元流域各个时刻的产流量rp1      
            open(17,file='./output/XAJ_RunoffDepth.txt')             
            do itime=1,Ntime
                write(17,2100)  (rp1(itime,isub),isub=1,nsub)
            enddo
            close(17)   
            
!-----------输出第二个微服务计算结果4：各个单元流域各个时刻的土壤含水量wp1      
            open(17,file='./output/XAJ_Soilmoisture_Whole.txt')             
            do itime=1,Ntime
                write(17,2100)  (wp1(itime,isub),isub=1,nsub)
            enddo
            close(17)                     

!-----------输出第二个微服务计算结果5：各个单元流域各个时刻的上层土壤含水量wup1      
            open(17,file='./output/XAJ_Soilmoisture_Upper.txt')             
            do itime=1,Ntime
                write(17,2100)  (wup1(itime,isub),isub=1,nsub)
            enddo
            close(17)                

!-----------输出第二个微服务计算结果6：各个单元流域各个时刻的上层土壤含水量wup1      
            open(17,file='./output/XAJ_Soilmoisture_Lower.txt')             
            do itime=1,Ntime
                write(17,2100)  (wlp1(itime,isub),isub=1,nsub)
            enddo
            close(17)    
            
    
!=======调用第三个微服务：新安江模型三水源划分微服务。根据净雨量pep1、产流量rp1、初始自由水蓄量S1、初始产流量FR1、参数KG1,KI1,IM1,SM1,EX1，计算时段长DeltaT,计算得到单元流域地表径流深rsp1、壤中流径流深rip1、地下径流深rgp1、更新的自由水蓄量sp1、更新的产流面积比frp1
     call XAJ3Div(pep1,rp1,S1,FR1,KG1,KI1,IM1,SM1,EX1,DeltaT,ntime,nsub,rsp1,rip1,rgp1,sp1,frp1)
     
!-----------输出第三个微服务计算结果1：各个单元流域各个时刻的地表径流深rsp1   
            open(17,file='./output/XAJ_Runoff_Surface.txt')             
            do itime=1,Ntime
                write(17,2101)  (rsp1(itime,isub),isub=1,nsub)
            enddo
            close(17)                
            
!-----------输出第三个微服务计算结果2：各个单元流域各个时刻的地表径流深rip1   
            open(17,file='./output/XAJ_Runoff_Inter.txt')             
            do itime=1,Ntime
                write(17,2101)  (rip1(itime,isub),isub=1,nsub)
            enddo
            close(17)               

 !-----------输出第三个微服务计算结果3：各个单元流域各个时刻的地表径流深rgp1   
            open(17,file='./output/XAJ_Runoff_Ground.txt')             
            do itime=1,Ntime
                write(17,2101)  (rgp1(itime,isub),isub=1,nsub)
            enddo
            close(17)    

 !-----------输出第三个微服务计算结果4：各个单元流域各个时刻的自由水蓄水量S1   
            open(17,file='./output/XAJ_Runoff_S.txt')             
            do itime=1,Ntime
                write(17,2100)  (sp1(itime,isub),isub=1,nsub)
            enddo
            close(17)      
            
!-----------输出第三个微服务计算结果5：各个单元流域各个时刻的产流面积比FR1   
            open(17,file='./output/XAJ_Runoff_FR.txt')             
            do itime=1,Ntime
                write(17,2101)  (frp1(itime,isub),isub=1,nsub)
            enddo
            close(17)    
            
2100        format(1x,50F10.2)  
2101        format(1x,50F14.4)  
            
!-----------调用第四个微服务：山坡汇流
            call  LagHillslopenew(CI1,CG1,Area1,DeltaT,ntime,nsub,rsp1,rip1,rgp1,QI1,QG1,qsp1,qip1,qgp1)
            
!-----------输出第四个微服务计算结果1：地表径流量qsp1
            open(17,file='./output/XAJ_Hillslope_QS.txt')             
            do itime=1,Ntime
                write(17,2101)  (qsp1(itime,isub),isub=1,nsub)
            enddo
            close(17)    
            
!-----------输出第四个微服务计算结果2：壤中流径流量qip1
            open(17,file='./output/XAJ_Hillslope_QI.txt')             
            do itime=1,Ntime
                write(17,2101)  (qip1(itime,isub),isub=1,nsub)
            enddo
            close(17)    
            
            
!-----------输出第四个微服务计算结果3：地下径流量qip1
            open(17,file='./output/XAJ_Hillslope_QG.txt')             
            do itime=1,Ntime
                write(17,2101)  (qgp1(itime,isub),isub=1,nsub)
            enddo
            close(17)                
            
 !-----------调用第五个微服务：河网汇流
            call  Lagnetworknew(CS1,L1,ntime,nsub,qsp1,qip1,qgp1,Qin1,Qinp1)
            
 !-----------输出第五个微服务计算结果1：河网汇流流量Qinp1   
            open(17,file='./output/XAJ_Network_Q.txt')             
            do itime=1,Ntime
                write(17,2101)  (Qinp1(itime,isub),isub=1,nsub)
            enddo
            close(17)                
 
            
2102        format(1x,F14.2)
            
 STOP     
End Program  MainMultisub
    