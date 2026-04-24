#include <stdio.h>
#define  M  5
void Calculate(int a[],int n,int b[])
{
int i,x,y,r;
for(i=0;i<n;i++)
{
x=a[i];
/***********SPACE***********/
y=a[【?】];(i+1)%n
do
{
/***********SPACE***********/
【?】;     r=x%y
x=y;
y=r;
}while(r);
b[i]=x;
}
}
void  main()
{
int  i,n=5,a[5]={18,66,38,87,15},b[5]={0};
Calculate(a,n,b);
for(i=0;i<n;i++)
printf("%3d、%3d的最大公约数：%3d\n",a[i],a[(i+1)%n],b[i]);
}