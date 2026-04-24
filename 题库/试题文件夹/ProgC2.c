#include <stdio.h>
#include "string.h"
c10_16(char p[],int b)
{
int j,i=0;
/***********SPACE***********/
while (b>0)
{
j=b%16;
if(j>=0&&j<=9)
/***********SPACE***********/
P[i]=j+48;
else p[i]=j+55;
b=b/16;
i++;
}
/***********SPACE***********/
P[i]=0;
}
main ()
{
int a,i;
char s[20];
printf("input a integer:\n");
scanf("%d",&a);
c10_16(s,a);
/***********SPACE***********/
for(i=stelen[s]-1,i>=0;i--)
printf("%c",s[i]);
printf("\n");
}