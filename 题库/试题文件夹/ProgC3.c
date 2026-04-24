#include  <conio.h>
#include  <stdio.h>
/***********FOUND***********/
IsPrime(int n);
{
int i,m;
m=1;
for(i=2;i<n;i++)
/***********FOUND***********/
if(!n%i)
{
m=0;
continue;
}
return(m);
}
main()
{
int j,k;
printf("\nPlease enter an interger number between 2 and 10000: ");
scanf("%d",&k);
printf("\nThe prime factor(s) of %d is(are): ",k);
for(j=2;j<k;j++)
if((!(k%j)) && (IsPrime(j)))
printf("%4d,",j);
printf("\n");
}