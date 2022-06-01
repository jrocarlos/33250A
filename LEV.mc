my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            33250A-3
DATE:                  2019-03-07 16:24:17
AUTHOR:                Carlos Júnior
REVISION:
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       1
NUMBER OF LINES:       121
=============================================================================
 STEP    FSC    RANGE NOMINAL        TOLERANCE     MOD1        MOD2  3  4 CON
  1.001  ASK-   R D   N B            P J S U       M C X Z        A  L  T  W
#------------CONFIG EXCEL---------------
  1.002  LIB          COM xlWS = xlApp.Worksheets["LEV"];
  1.003  LIB          xlWS.Select();
#---------------------------PERDA POR INSERÇÃO----------------------
  1.004  MEMI         INSIRA O VALOR DE PERDA DO CONECTOR
  1.005  MATH         MEM2 = MEM
#----------------------------CONFIG GENERATOR----------------------
  1.006  RSLT         =
  1.007  IEEE         [@13]*RST
  1.008  IEEE         *CLS
  1.009  IEEE         :VOLTage:UNIT DBM
#---------------------------CONFIG METER-----------------------
  1.010  IEEE         [@21]*RST
  1.011  IEEE         SYST:PRES
  #1.017  IEEE         CAL:ZERO:AUTO ONCE
  #1.018  IEEE         CAL:ALL:ZERO:FAST:AUTO
  #1.019  IEEE         SYST:ERR?
  #1.020  IEEE         SENS:AVER:COUN:AUTO ON
  #1.021  IEEE         SENS:AVER:COUN 16
  #1.022  IEEE         OUTP:ROSC ON
#-----------------ZERO METER----------------
  1.012  DISP         CONECTE O SENSOR NA PORTA "POWER REF"
  1.012  DISP
  1.012  DISP         [32]   POWER METER         to         SENSOR
  1.012  DISP         [32]
  1.012  DISP         [32]
  1.012  DISP         [32]   POWER REF -------------------> SENSOR
  1.012  DISP         [32]
  1.012  DISP         [32]     GPIB POWER METER NRP2 = 21
  1.012  DISP         [32]     GPIB GERADOR 33120A = 13
  1.012  DISP         [32]
  1.013  PIC          SETUP2-1
  1.014  IEEE         SENS1:FREQ:CW 50 MHZ
  1.015  IEEE         INIT:CONT ON
  1.016  IEEE         CAL1:ZERO:AUTO ONCE
  1.017  WAIT         [D7000]
  1.018  IEEE         OUTP:ROSC ON
  1.019  WAIT         [D8000]
  1.020  IEEE         OUTP:ROSC OFF
  1.021  DISP         Connect the generator to the UUT as follows:
  1.021  DISP
  1.021  DISP         [32]   Generator         to         Meter
  1.021  DISP         [32]   OUTPUT -------------------> POWER SENSOR
  1.021  DISP         [32]
  1.022  PIC          SETUP2-2
#-------------------CONFIG  Nº MEAS----------------
  1.023  MEMI         DIGITE O NÚMERO DE MEDIDAS
  1.024  MATH         A = MEM
#-----------------CONFIG POINT------------------
  1.025  MATH         P = 0
  1.026  MATH         LP = 2
  1.027  MATH         CP = 1
  1.028  MATH         T  = 0
  1.029  MATH         L = 0
  1.030  MATH         LINHA = 2
  1.031  MATH         COLUNA = 5
  1.032  MATH         CPERDA = COLUNA + A
  1.033  DO
  1.034  LIB          COM P1 = xlApp.Cells[LP,CP];
  1.035  LIB          PONTO = P1.Value2;
  1.036  IF           PONTO == 0
  1.037  JMP          1.076
  1.038  ENDIF
  1.039  MATH         CP = CP + 1
  1.040  LIB          COM T1 = xlApp.Cells[LP,CP];
  1.041  LIB          TEX = T1.Value2;
  1.042  MATH         P = PONTO&TEX
#----------------LEVEL-------------------
  1.043  MATH         CP = CP + 1
  1.044  LIB          COM L1 = xlApp.Cells[LP,CP];
  1.045  LIB          L = L1.Value2;
#----------------------END-------------------------------
  1.046  IF           P == 00
  1.047  JMP          1.076
  1.048  ENDIF
#----------------------------CONFIG OUT GENERATOR--------------
  1.049  IEEE         [@13]:FREQ [V P]
  1.050  IEEE         :VOLT:LEV [V L]
  1.051  IEEE         OUTP ON
  1.052  WAIT         [D2000]
#------------CONFIG IN COUNT----------------
  1.053  MATH         TEMPO = 5
  1.054  IEEE         [@21]SENS1:FREQ:CW  [V P]
  1.055  IEEE         INIT:CONT ON
  1.056  DO
  1.057  WAIT         -t [V TEMPO] Please Standby
  1.058  IEEE         READ?[I]
#------------------SAVE DATE----------------
  1.059  LIB          COM selectedCell = xlApp.Cells[LINHA,COLUNA];
  1.060  LIB          selectedCell.Select();
  1.061  LIB          selectedCell.FormulaR1C1 = [MEM];
  1.062  MATH         T = T + 1
  1.063  MATH         COLUNA = COLUNA + 1
  1.064  MATH         CP = CP + 1
  1.065  UNTIL        T == A
#---------------------SAVE LOSS---------------------
  1.066  LIB          COM selectedCell2 = xlApp.Cells[LINHA,CPERDA];
  1.067  LIB          selectedCell2.Select();
  1.068  LIB          selectedCell2.Value2 = [MEM2];
  1.069  MATH         T  = 0
  1.070  MATH         COLUNA = 5
  1.071  MATH         LINHA = LINHA + 1
  1.072  MATH         CP = 1
  1.073  MATH         LP = LP + 1
  1.074  UNTIL        PONTO == 0
  1.075  JMP          1.076
#------------------RESET------------------
  1.076  IEEE         [@21]*RST
  1.077  IEEE         [@13]*RST
