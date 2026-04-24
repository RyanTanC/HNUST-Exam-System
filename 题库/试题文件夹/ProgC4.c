#include   <conio.h>
#include   <stdio.h>
#define    N   20
/***********FOUND***********/
void fun(int a[],int n);
{
int  i,  m, t, k;
for(i=0; i<2;i++)
{
/***********FOUND***********/
m=i;
for(k=i+1; k<n; k++)
if(a[k]>a[m])
m=k;
t=a[i];
a[i]=a[m];
a[m]=t;
}
}
main()
{
int b[N]={11,5,12,0,3,6,9,7,10,8},n=10,i;
system("CLS");
for(i=0; i<n; i++)
printf("%d ",b[i]);
printf("\n");
fun(b,n);
for(i=0; i<n; i++)
printf("%d ", b[i]);
printf("\n");
}