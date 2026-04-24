#include    <stdio.h>
#define    N    9
int fun(int  a[], int  n)
{
int  i,j;
j = 0;
for (i=0; i<n; i++)
/***********SPACE***********/
if (a[i]%2==【?】)       1
{
/***********SPACE***********/
a[j] = a[i]; 【?】;       j++
}
/***********SPACE***********/
return 【?】;       j
}
main()
{
int  b[N]={9,1,4,2,3,6,5,8,7}, i, n;
printf("\nThe original data  :\n");
for (i=0; i<N; i++)  printf("%4d ", b[i]);
printf("\n");
n = fun(b, N);
printf("\nThe number of odd  : %d \n", n);
printf("\nThe odd number  :\n");
for (i=0; i<n; i++)  printf("%4d ", b[i]);
printf("\n");
}