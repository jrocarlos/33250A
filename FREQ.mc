my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            33250A-2
DATE:                  2019-03-07 16:24:16
AUTHOR:                Carlos Júnior
REVISION:
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       2
NUMBER OF LINES:       138
=============================================================================
 STEP    FSC    RANGE NOMINAL        TOLERANCE     MOD1        MOD2  3  4 CON
  1.001  DISP         Connect the generator to the UUT as follows:
  1.001  DISP
  1.001  DISP         [32]   Generator         to         Counter
  1.001  DISP         [32]     OUTPUT -------------------> CHANNEL 1
  1.001  DISP         [32]
  1.001  DISP         [32]     GPIB CONTADOR 53132A = 3
  1.001  DISP         [32]     GPIB GERADOR 33521A = 13
  1.002  PIC          SETUP1
  1.003  ASK-   R D   N B            P J S U       M C X Z        A  L  T  W
#------------CONFIG EXCEL---------------
  1.004  LIB          COM xlWS = xlApp.Worksheets["FREQ"];
  1.005  LIB          xlWS.Select();
#------------------CONFIG GENERATOR---------
  1.006  TARGET       -p
  1.007  RSLT         =
  1.008  IEEE         [@13]*RST
  1.009  IEEE         *CLS
 # IEEE         :VOLTage:UNIT DBM
  1.010  IEEE         :VOLTage:UNIT VRMS
  1.011  IEEE         :VOLT:LEV 1
#-----------------CONFIG COUNT----------------
  1.012  IEEE         [@3]*RST
  1.013  IEEE         :FUNC 'FREQ 1'
  1.014  IEEE         INIT:CONT OFF
  1.015  IEEE         INP1:COUP DC
  1.016  IEEE         INP1:IMP 50
  1.017  IEEE         EVEN1:LEV 0.5
  1.018  TARGET       -m
#-------------------CONFIG  Nº MEAS----------------
  1.019  MEMI         DIGITE O NÚMERO DE MEDIDAS
  1.020  MATH         A = MEM
#-----------------CONFIG POINT------------------
  1.021  MATH         P = 0
  1.022  MATH         LP = 2
  1.023  MATH         CP = 1
  1.024  MATH         T  = 0
  1.025  MATH         LINHA = 2
  1.026  MATH         COLUNA = 3
  1.027  DO
  1.028  LIB          COM P1 = xlApp.Cells[LP,CP];
  1.029  LIB          PONTO = P1.Value2;
  1.030  IF           PONTO == 0
  1.031  JMP          2.019
  1.032  ENDIF
  1.033  MATH         CP = CP + 1
  1.034  LIB          COM T1 = xlApp.Cells[LP,CP];
  1.035  LIB          TEX = T1.Value2;
  1.036  MATH         P = PONTO&TEX
  1.037  MATH         EX = 0
  1.038  MATH         Z1 = CMP  (TEX,"MHz")
  1.039  MATH         Z2 = CMP  (TEX,"kHz")
  1.040  MATH         Z3 = CMP  (TEX,"Hz")
#-------------------TRIGGER---------------------------
  1.041  IF           PONTO > 10 && Z3 == 1
  1.042  JMP          2.015
  1.043  ENDIF
#-------------------FILTER---------------------------
  1.044  IF           PONTO > 100 && Z2 == 1 || Z1 == 1
  1.045  JMP          2.017
  1.046  ELSE
  1.047  IEEE         INP1:FILT ON
  1.048  ENDIF
#----------------------END-------------------------------
  1.049  IF           P == 00
  1.050  JMP          2.019
  1.051  ENDIF
#----------------------GATE-------------------------------
  1.052  IF           PONTO < 1 && Z3 == 1
  1.053  MATH         GATE = 100
  1.054  ELSE
  1.055  MATH         GATE = 10
  1.056  ENDIF
#----------------------------CONFIG OUT GENERATOR--------------
  1.057  IEEE         [@13]:FREQ [V P]
  1.058  IEEE         OUTP ON
  1.059  WAIT         [D2000]
#------------CONFIG IN COUNT----------------
  1.060  MATH         TEMPO = GATE + (GATE / 2)
  1.061  IEEE         [@3]FREQ:ARM:STOP:TIM [V GATE]
  1.062  TARGET       -m
  1.063  IEEE         INIT:CONT ON
  1.064  DO
  1.065  WAIT         -t [V TEMPO] Please Standby
  1.066  IEEE         READ:FREQ?[I]
  1.067  IF           Z1 == 1
  1.068  MATH         EX = 1E6
  1.069  ENDIF
  1.070  IF           Z2 == 1
  1.071  MATH         EX = 1E3
  1.072  ENDIF
  1.073  IF           Z3 == 1
  1.074  MATH         EX = 1E0
  1.075  ENDIF
  1.076  IF           Z1 == 1 && PONTO == 1
  1.077  MATH         EX = 1E6
  1.078  ENDIF
  1.079  IF           Z2 == 1 && PONTO == 1
  1.080  MATH         EX = 1E3
  1.081  ENDIF
  1.082  IF           Z3 == 1 && PONTO == 1
  1.083  MATH         EX = 1E0
  1.084  ENDIF
  1.085  MATH         MEM = MEM / EX
  1.086  MEMCX        0              TOL
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
  2.014  JMP          2.019
#----------------CONFIG TRIGGER-------------------
  2.015  IEEE         EVEN1:LEV:AUTO ON
  2.016  JMP          1.049
#----------------CONFIG FILTER 100k-------------------
  2.017  IEEE         INP1:FILT OFF
  2.018  JMP          1.048
#------------------RESET------------------
  2.019  IEEE         [@3]*RST
  2.020  IEEE         [@13]*RST
