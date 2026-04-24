#include   <stdio.h>
#include   <math.h>
double fun(double  x)
{
double  f, t;      int  n;
f = 1.0 + x;
/***********SPACE***********/
t=【?】;
n = 1;
do
{
n++;
/***********SPACE***********/
t*=(-1.0)*x/【?】;
f += t;
}
/***********SPACE***********/
while(【?】 >=1e-6);
return  f;
}
main()
{
double x, y;
x=2.5;
y = fun(x);
printf("\nThe result is :\n");
printf("x=%-12.6f y=%-12.6f\n", x, y);
}