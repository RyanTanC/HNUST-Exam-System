#include<stdio.h>
main()
{
int m,n;
float x,term,ex1,ex2;
printf("x,m=");
scanf("%f %d",&x,&m);
/**********FOUND**********/
ex1==ex2=1;
term=1;
for(n=1;n<=m;n++)
{
/**********FOUND**********/
term*=x%n;
ex1+=term;
}
ex2=term;
/**********FOUND**********/
for (n=m;n>1; n--)
{
term*=n/x;
ex2+=term;
}
printf("exforward=%f exbackrard=%f\n",ex1,ex2);
}