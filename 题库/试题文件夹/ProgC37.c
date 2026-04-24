#include <stdio.h>
double fun(int m)
{
/**********FOUND**********/
double y=0
int i;
/**********FOUND**********/
for(i=1; i<m; i++)
{
/**********FOUND**********/
y=+1.0/(2*i*i);
}
return(y);
}
main()
{
int n;
printf("Enter n: ");
scanf("%d", &n);
printf("\nThe result is %1f\n", fun(n));
}