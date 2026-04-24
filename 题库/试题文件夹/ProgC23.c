#include  <stdio.h>
int  cyclic_min(int  x[], int n)
{
int  left  = 0;
int  right = n - 1;
int  mid;
/***********SPACE***********/
while (【?】)        right>left
{
mid = (left + right)/2;
if (x[mid] < x[right])
/***********SPACE***********/
【?】;                        left=mid
else
/***********SPACE***********/
【?】;                        left=mid+1
}
return left;
}
#include  <stdio.h>
void  main(void)
{
int  x[] = { 20, 23, 28, 35, 39, 40, 42, 8, 10, 15, 17, 19};
int  n   = sizeof(x)/sizeof(int);
int  loc, i;
printf("\nFind Cyclic Minimum");
printf("\n===================");
printf("\n\nGiven Array Sorted in Cyclic Fashion :\n");
for (i = 0; i < n; i++)
printf("%3d", x[i]);
loc = cyclic_min(x, n);
printf("\n\nMinimum is located at x[%d] = %d", loc, x[loc]);
}