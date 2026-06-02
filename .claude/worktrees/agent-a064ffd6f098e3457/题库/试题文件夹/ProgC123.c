/*-------------------------------------------------------
【程序改错】
---------------------------------------------------------
题目：下列给定程序中，fun函数的功能是：根据形参m，计算下列公式的值。
       t＝1＋1/2＋1/3＋1/4＋…＋1/m
例如：若输入5，则应输出2.283333。
-------------------------------------------------------*/
#include <stdlib.h>
#include <conio.h>
#include <stdio.h>
double fun(int m)
{
        double t=1.0;
        int i;
        for(i=2;i<=m;i++)
/***********FOUND***********/
                t+=1.0/k;             /*只修改错误的地方其他不要改写否则不得分*/
/***********FOUND***********/
        return i;
}
void main()
{
        int m;
        system("CLS");
        printf("\nPlease enter 1integer number: ");
        scanf("%d",&m);
        printf("\nThe result is %1f\n", fun(m));
}