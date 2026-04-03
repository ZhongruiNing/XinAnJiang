        !COMPILER-GENERATED INTERFACE MODULE: Sat Sep 18 17:36:22 2021
        MODULE XAJ3DIV__genmod
          INTERFACE 
            SUBROUTINE XAJ3DIV(PEP1,RP1,S1,FR1,KG1,KI1,IM1,SM1,EX1,     &
     &DELTAT,NTIME,NSUB,RSP1,RIP1,RGP1,SP1,FRP1)
              INTEGER(KIND=4) :: NSUB
              INTEGER(KIND=4) :: NTIME
              REAL(KIND=4) :: PEP1(NTIME,NSUB)
              REAL(KIND=4) :: RP1(NTIME,NSUB)
              REAL(KIND=4) :: S1(NSUB)
              REAL(KIND=4) :: FR1(NSUB)
              REAL(KIND=4) :: KG1(NSUB)
              REAL(KIND=4) :: KI1(NSUB)
              REAL(KIND=4) :: IM1(NSUB)
              REAL(KIND=4) :: SM1(NSUB)
              REAL(KIND=4) :: EX1(NSUB)
              REAL(KIND=4) :: DELTAT
              REAL(KIND=4) :: RSP1(NTIME,NSUB)
              REAL(KIND=4) :: RIP1(NTIME,NSUB)
              REAL(KIND=4) :: RGP1(NTIME,NSUB)
              REAL(KIND=4) :: SP1(NTIME,NSUB)
              REAL(KIND=4) :: FRP1(NTIME,NSUB)
            END SUBROUTINE XAJ3DIV
          END INTERFACE 
        END MODULE XAJ3DIV__genmod
