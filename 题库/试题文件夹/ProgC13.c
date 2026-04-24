#include<stdio.h>
void sort(int [ ],int );
int main()
{
int  i, a[ ]={6,0,-3,9,4,-5,18,20};
int n=sizeof(a)/sizeof(a[0]);
/***********FOUND***********/
sort(n ,a[n]);
for(i=0;i<n;i++)
printf("%d\t",a[i]);
printf("\n");
return 0;
}
void sort(int a[ ],int n)
{
int i,j,k,swap;
for(i=0;i<n-1;i++)
{
swap=0;
/***********FOUND***********/
for(j=n-1;j>i;j++)
if(a[j]>a[j-1])
k=a[j],a[j]=a[j-1],a[j-1]=k,swap=1;
/***********FOUND***********/
if(swap)break;
}
}