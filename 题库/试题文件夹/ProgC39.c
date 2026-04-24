#include <stdio.h>
main()
{
char c;
/***********SPACE***********/
while((c=【?】)!='\n')
{
/***********SPACE***********/
if((c>='a'&&c<='z')||(c>='A'&&c<='Z'))【?】;
/***********SPACE***********/
if((c>'Z'【?】c<='Z'+4)||c>'z') c-=26;
printf("%c",c);
}
}