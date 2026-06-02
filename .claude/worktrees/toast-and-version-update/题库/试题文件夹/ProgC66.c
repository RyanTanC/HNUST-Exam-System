// 功能：下面的程序是求 1!+3!+5!+……+n! 的和。
#include <stdio.h>
main(){
long int f,s;
int i,j,n;
【1】s=0;
scanf("%d",&n);
for(i=1;i<=n;【2】i+=2) {
f=1;
for(j=1;【3】j<=i;j++)
【4】f=f*j;
s=s+f;
}
printf("n=%d,s=%ld\n",n,s);
}