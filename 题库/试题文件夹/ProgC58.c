#include<stdio.h>
#include<string.h>
void main()
{
int fun(char s[]);
char s[80];
int c;
printf("Please input the string:");
gets(s);
c=fun(s);
printf("c=%d\n",c);
}
int fun(char s[])
{
int c1,c2,i;
i=c1=c2=0;
while(s[i]!='\0')
{
if(s[i]=='(') c1++;
else if(s[i]==')')
{
c2++;
/***********SPACE***********/
if(【?】) return(0);
}
i++;
}
/***********SPACE***********/
return(【?】);
}