
typedef union PORTu{

	struct{
		unsigned char byte;
		unsigned char ID;
	};
	struct{
      unsigned CON0:1;
      unsigned CON1:1;
      unsigned CON2:1;
      unsigned CON3:1;
      unsigned CON4:1;
      unsigned CON5:1;
      unsigned CON6:1;
      unsigned CON7:1;

  };
	
}PORT;
void Setpin(PORT *P1,char channel,char value);
void Writeport(PORT *P1);
void DigitalWrite(PORT *P1,char channel,char value);
char DigitalRead(PORT *P1,char channel);
char Readpin(PORT *P1,char channel);
char Readport(PORT *P1);
void pinmode(PORT *P1,char channel,char value);
void pinmode(PORT *P1,char channel,char value)
{
		
}

void DigitalWrite(PORT *P1,char channel,char value)
{		
			Setpin(P1,channel,value);
			Writeport(P1);
}
void Setpin(PORT *P1,char channel,char value)
{
			if(value)
				P1->byte|=1<<channel;
			else
				P1->byte&=~(1<<channel);

}
void Writeport(PORT *P1)
{		
		switch(P1->ID)
		{
			case 1:
				LATA=P1->byte;

				break;
			case 2:
				LATB=P1->byte;

				break;
			case 3:
				LATC=P1->byte;

				break;
			case 4:
				LATD=P1->byte;

				break;
			case 5:
				LATE=P1->byte;

				break;
		}

}

char DigitalRead(PORT *P1,char channel)
{		
		Readport(P1);
		Readpin(P1,channel);
}
char Readpin(PORT *P1,char channel)
{
			if(P1->byte&(1<<channel))
				return 1;
			else 
				return 0;
}
char Readport(PORT *P1)
{		

		switch(P1->ID)
		{
			case 1:
				P1->byte=PORTA;
				break;
			case 2:
				P1->byte=PORTB;
				break;
			case 3:
				P1->byte=PORTC;
				break;
			case 4:
				P1->byte=PORTD;
				break;
			case 5:
				P1->byte=PORTE;
				break;
		}
}
