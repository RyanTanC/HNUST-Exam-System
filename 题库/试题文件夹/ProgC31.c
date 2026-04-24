#include <stdio.h>
long fun(int n)
{  int i;
/**********FOUND**********/
int s=0;
for(i=1;i<n;i++)
/**********FOUND**********/
if(i%3=0)
s+=i;
return s;
}
void main()
{
int n;
long int  result;
printf("Enter n: ");
scanf("%d",&n);
result=fun(n);
printf("Result=%ld\n",result);
}