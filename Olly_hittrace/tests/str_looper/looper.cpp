// looper.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"

char* decode(char* s, int len, char key){
	for(int i=0;i<len;i++) s[i] = s[i] ^ key;
	return s;
}


void main(void)
{

	char testing[8] = {0x21, 0x30, 0x26, 0x21, 0x3C, 0x3B, 0x32,0x00};
	char blahblah[9] = {0x15, 0x1B, 0x16, 0x1F, 0x15, 0x1B, 0x16, 0x1F,0x00};

	decode(testing, 7, 0x55);
	decode(blahblah, 8, 0x77);

	printf("String 1 is now: %s\n", testing);
	printf("String 2 is now: %s\n", blahblah);

}



