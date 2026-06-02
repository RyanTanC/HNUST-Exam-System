/*------------------------------------------------------    
【程序改错】
--------------------------------------------------------
功能：输入一个字符串，过滤此串，滤掉字母字符，并统计新生
      成串中包含的字符个数。
例如：输入的字符串为ab234$df4，则输出为：
      The new string is 234$4 
      There are 5 char in the new string.。
------------------------------------------------------*/
#include <stdio.h>
#include <conio.h>
#define N 80
int fun(char *ptr)
{
  int i,j;
  /**********FOUND**********/
  for(i=0,j=0;*(ptr+i)!="\\0";i++)
    /**********FOUND**********/
    if(*(ptr+i)>'z'|| *(ptr+i)<'a'||*(ptr+i)>'Z' || *(ptr+i)<'A')
    {
      /**********FOUND**********/
      (ptr+j)=(ptr+i);
      j++;
    }
  *(ptr+j)='\0';
  return(j);
}
main()
{
  char str[N];
  int s;
  printf("input a string:");gets(str);
  printf("The original string is :"); puts(str);
  s=fun(str);
  printf("The new string is :");puts(str);
  printf("There are %d char in the new string.",s);
}