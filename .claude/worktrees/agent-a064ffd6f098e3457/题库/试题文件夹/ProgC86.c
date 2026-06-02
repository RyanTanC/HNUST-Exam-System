/*-------------------------------------------------------
【程序设计】
---------------------------------------------------------
题目：编写函数fun，w是一个大于10的无符号整数，若w是n(n≥2)位的整数，则函数求
      出w的后n－1位的数作为函数值返回。
例如：w值为5923，则函数返回923；若w值为923，则函数返回23。
注意：请勿改动main函数和其他函数中的任何内容，仅在函数fun的花括号中填入
      你编写的若干语句。
-------------------------------------------------------*/
#include<conio.h>
#include<stdio.h>
#include<stdlib.h>
unsigned fun(unsigned w)
{
/**********Program**********/
/**********  End  **********/
}
void main()
{ 
        unsigned x;
        printf("Enter a unsigned integer number: ");
        scanf ("%u",&x);
        printf("The original data is:%u\n",x);
        if(x<10) 
                printf("Data error! ");
        else 
                printf ("The result :%u\n", fun(x));
}