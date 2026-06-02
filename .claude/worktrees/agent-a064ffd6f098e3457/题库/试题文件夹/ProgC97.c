/*------------------------------------------------------        
【程序改错】
--------------------------------------------------------
功能：在一个一维整型数组中找出其中最大的数及其下标。
------------------------------------------------------*/
#include <stdio.h>
#define N 10
/**********FOUND**********/
float fun(int *a,int *b,int n)
{
  int *c,max=*a;
  for(c=a+1;c<a+n;c++)
    if(*c>max)
    {
      max=*c;
      /**********FOUND**********/
      b=c-a;
    }
  return max;
}
void main()
{
  int a[N],i,max,p=0;
  printf("please enter 10 integers:\n");
  for(i=0;i<N;i++)
    /**********FOUND**********/
    get("%d",a[i]);
  /**********FOUND**********/
  m=fun(a,p,N);
  printf("max=%d,position=%d",max,p);
}