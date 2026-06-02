/*-------------------------------------------------------
【程序改错】
---------------------------------------------------------
题目：下列给定程序中函数fun的功能是：将p所指字符串中的所有字符复制到b中，要求
      每复制三个字符之后插入一个空格。
例如：若给a输入字符串：ABCDEFGKHIJK，调用函数后，字符数组b中的内容
      为：ABC　DEF　GHI　JK。
-------------------------------------------------------*/
#include <stdio.h>
void  fun(char  *p, char  *b)
{  
        int   i, k=0;
        while(*p)
        {  
                i=1;
                while( i<=3 && *p ) 
                {
        /***********FOUND***********/
                        b[k]=p;
                        k++; p++; i++;
                }
                if(*p)
                {
        /***********FOUND***********/
                        b[k++]=" ";
                }
        }
        b[k]='\0';
}
main()
{  
        char  a[80],b[80];
        printf("Enter a string:      "); 
        gets(a);
        printf("The original string: "); 
        puts(a);
        fun(a,b);
        printf("\nThe string after insert space:   "); 
        puts(b); printf("\n\n");
}