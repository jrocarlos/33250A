my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            33250A-4
DATE:                  2019-03-07 16:24:17
AUTHOR:                Carlos Júnior
REVISION:
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       2
NUMBER OF LINES:       108
=============================================================================
 STEP    FSC    RANGE NOMINAL        TOLERANCE     MOD1        MOD2  3  4 CON
  1.001  DISP
  1.001  DISP         [32]   Generator         to         Counter
  1.001  DISP         [32]   REFERENCE -------------------> CHANNEL 1
  1.001  DISP         [32]
  1.001  DISP         [32]     GPIB CONTADOR 53132A = 3
  1.002  PIC          SETUP3
  1.003  ASK-   R D   N B            P J S U       M C X Z        A  L  T  W
#-------------------CONFIG EXCEL-----------------------
  1.004  LIB          COM xlWS = xlApp.Worksheets["REF"];
  1.005  LIB          xlWS.Select();
#-----------------CONFIG COUNT----------------
  1.006  IEEE         [@3]*RST
  1.007  IEEE         :FUNC 'FREQ 1'
  1.008  IEEE         INIT:CONT OFF
  1.009  IEEE         INP1:COUP DC
  1.010  IEEE         INP1:IMP 50
  1.011  IEEE         EVEN1:LEV:AUTO ON
  1.012  IEEE         INP1:FILT OFF
  1.013  TARGET       -m
#-------------------CONFIG  Nº MEAS----------------
  1.014  MEMI         DIGITE O NÚMERO DE MEDIDAS
  1.015  MATH         A = MEM
#-----------------CONFIG POINT------------------
  1.016  MATH         P = 0
  1.017  MATH         LP = 2
  1.018  MATH         CP = 1
  1.019  MATH         T  = 0
  1.020  MATH         LINHA = 2
  1.021  MATH         COLUNA = 3
  1.022  DO
  1.023  LIB          COM P1 = xlApp.Cells[LP,CP];
  1.024  LIB          PONTO = P1.Value2;
  1.025  IF           PONTO == 0
  1.026  JMP          2.014
  1.027  ENDIF
  1.028  MATH         CP = CP + 1
  1.029  LIB          COM T1 = xlApp.Cells[LP,CP];
  1.030  LIB          TEX = T1.Value2;
  1.031  MATH         P = PONTO&TEX
  1.032  MATH         EX = 0
  1.033  MATH         Z1 = CMP  (TEX,"MHz")
  1.034  MATH         Z2 = CMP  (TEX,"kHz")
  1.035  MATH         Z3 = CMP  (TEX,"Hz")
#----------------------END-------------------------------
  1.036  IF           P == 00
  1.037  JMP          2.014
  1.038  ENDIF
#----------------------END-------------------------------
  1.039  IF           PONTO < 1 && Z3 == 1
  1.040  MATH         GATE = 100
  1.041  ELSE
  1.042  MATH         GATE = 10
  1.043  ENDIF
#------------CONFIG IN COUNT----------------
  1.044  MATH         TEMPO = GATE + (GATE / 2)
  1.045  IEEE         [@3]FREQ:ARM:STOP:TIM [V GATE]
  1.046  TARGET       -m
  1.047  IEEE         INIT:CONT ON
  1.048  DO
  1.049  WAIT         -t [V TEMPO] Please Standby
  1.050  IEEE         READ:FREQ?[I]
  1.051  IF           Z1 == 1
  1.052  MATH         EX = 1E6
  1.053  ENDIF
  1.054  IF           Z2 == 1
  1.055  MATH         EX = 1E3
  1.056  ENDIF
  1.057  IF           Z3 == 1
  1.058  MATH         EX = 1E0
  1.059  ENDIF
  1.060  IF           Z1 == 1 && PONTO == 1
  1.061  MATH         EX = 1E6
  1.062  ENDIF
  1.063  IF           Z2 == 1 && PONTO == 1
  1.064  MATH         EX = 1E3
  1.065  ENDIF
  1.066  IF           Z3 == 1 && PONTO == 1
  1.067  MATH         EX = 1E0
  1.068  ENDIF
  1.069  MATH         MEM = MEM / EX
  1.070  MEMCX        0              TOL
  2.001  LIB          COM selectedCell = xlApp.Cells[LINHA,COLUNA];
  2.002  LIB          selectedCell.Select();
  2.003  LIB          selectedCell.FormulaR1C1 = [MEM];
  2.004  MATH         T = T + 1
  2.005  MATH         COLUNA = COLUNA + 1
  2.006  MATH         CP = CP + 1
  2.007  UNTIL        T == A
  2.008  MATH         T  = 0
  2.009  MATH         COLUNA = 3
  2.010  MATH         LINHA = LINHA + 1
  2.011  MATH         CP = 1
  2.012  MATH         LP = LP + 1
  2.013  UNTIL        PONTO == 0
#------------------RESET------------------
  2.014  IEEE         [@3]*RST
  2.015  IEEE         [@13]*RST
