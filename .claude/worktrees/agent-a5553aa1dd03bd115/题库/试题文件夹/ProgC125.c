/*--------------------------------------------------------------------
【程序设计】
----------------------------------------------------------------------
题目：编写函数int isPalindrome(const char str[], int begin, int end)，
      判断一个字符串是否为回文字符串。
---------------------------------------------------------------------*/
#include <stdio.h>
#include <string.h>
int isPalindrome(const char str[], int begin, int end)
{
/**********Program**********/
/**********  End  **********/
}
int main()
{
    char a[100]= {0};
    gets(a);
    if(isPalindrome(a,0,strlen(a)-1))
        printf("%s是回文字符串\n",a);
    else
        printf("%s不是回文字符串\n",a);
    return 0;
}