#include <stdio.h>
#include <conio.h>
/**********FOUND**********/
#define N= 7      宏定义
main()
{
char a[N][N];
int i,j,z;
for(i=0;i<N;i++)
for(j=0;j<N;j++)
/**********FOUND**********/
a[i][j]=;       ‘ ’
z=0;
for(i=0;i<(N+1)/2;i++)
{
for(j=z;j<N-z;j++)
a[i][j]='*';
z=z+1;
}
/**********FOUND**********/
z=0;
for(i=(N+1)/2;i<N;i++)
{
z=z-1;
for(j=z;j<N-z;j++)
a[i][j]='*';
}
for(i=0;i<N;i++)
{
for(j=0;j<N;j++)
/**********FOUND**********/
printf("%d",a[i][j]);        c
printf("\n");
}
}