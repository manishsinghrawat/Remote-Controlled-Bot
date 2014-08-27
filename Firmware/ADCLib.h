#define Max_Channel 13
#include <adc.h>
#include<delays.h>
void Adc_Init();
void Adc_Init()
{
   //We use default value for +/- Vref

   //VCFG0=0,VCFG1=0
   //That means +Vref = Vdd (5v) and -Vref=GND
   // ADCON1 is left default.

   //Port Configuration
   //We also use default value here too
   //All ANx channels are Analog

   /*
      ADCON2

      *ADC Result Right Justified.
      *Acquisition Time = 2TAD
      *Conversion Clock = 32 Tosc
   */
//	ADCON1 = 0b00001110;               //Set all pins to digital I/O except RA0/AN0. (page 262 of datasheet)
//	TRISA  = 0b00000001;               // Set I/0 for PORTA
unsigned char i;
OpenADC( ADC_FOSC_32      &

           ADC_RIGHT_JUST   &

           ADC_20_TAD,

           ADC_CH0          &

                     ADC_REF_VDD_VSS  &

           ADC_INT_OFF,ADC_CH0);

 //  ADCON2=0b10001010;
for( i=0;i<255;i++);
}



unsigned int AnalogRead(unsigned char channel)
{
   	if(channel>Max_Channel) return 0;  //Invalid Channel
   	//SelChanConvADC(0b10000111|(channel<<3));
 	
	SetChanADC(0b10000111|(channel<<3));
  Delay10TCYx(255);
	ConvertADC();
	while ( BusyADC() );
	return ReadADC();
}



    

