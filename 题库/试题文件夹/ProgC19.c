#include <stdlib.h>
#include <conio.h>
#include <stdio.h>
double  fun(int   x[ ])
{
double sum = 0.0;
int  c = 0, i = 0;
/***********FOUND***********/
while(x[i]==0)
{
if(x[i]<0)
{
sum=sum+x[i];
c++;
}
i++;
}
/***********FOUND***********/
sum=sum\c;
return  sum;
}
void main()
{
int  x[1000];
int  i=0;
system("CLS");
printf("\nPlease enter some data(end with 0) :");
do
{
scanf("%d",&x[i]);
}
while(x[i++]!=0);
printf("%f\n",fun(x));
}