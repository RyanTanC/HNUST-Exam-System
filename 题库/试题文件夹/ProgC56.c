#include <stdio.h>
#define N 11
void main()
{
int a[N],x,p;
int i;
printf("Please input %d numbers:",N-1);
for(i=0;i<=N-2;i++)
scanf("%d",&a[i]);
printf("Please input x to be intert:");
scanf("%d",&x);
/**********Program**********/
for(i=0;i<=N-1;i++)
{if(a[i]>=x)
{p=I;
Break;}
}
for(i=N-2;i>=p;i--)
a[i+1]=a[i];
a[p]=x;
/**********  End  **********/
for(i=0;i<=N-1;i++)//结果
printf("%5d",a[i]);
printf("\n");
}