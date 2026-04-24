#include <stdio.h>
int fun(int k)
{
int m=0,mc=0;
/**********FOUND**********/
while ((k>=2)||(mc<10))
{
/**********FOUND**********/
if((k%13=0)||(k%17=0))
{
m=m+k;
mc++;
}
/**********FOUND**********/
k++;
}
/**********FOUND**********/
return  ;
}
main()
{
printf("%d\n",fun(500));
}