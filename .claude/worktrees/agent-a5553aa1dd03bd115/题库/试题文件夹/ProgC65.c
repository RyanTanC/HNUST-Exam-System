// 功能：将一个字符串中下标为 m 的字符开始的全部字符复制成为另一个字符串。
#include<stdio.h>
void strcopy(char *str1,char *str2,int m)
{
    char *p1,*p2;
    p1 = str1 + m;
    p2 = str2;
    while(*p1)
        *p2++ = *p1++;
    *p2 = '\0';
}
main()
{
    int m;
    char str1[80],str2[80];
    gets(str1);
    scanf("%d",&m);
    strcopy(str1, str2, m);
    puts(str1);
    puts(str2);
}