/*-------------------------------------------------------
【程序改错】
---------------------------------------------------------
函数max4()和max2()分别求出4个整数和2个整数的最大值。
下面给定的程序存在错误，请改正。
注意：不得增行或删行，也不得更改程序的结构。
-------------------------------------------------------*/
#include <stdio.h>
int max2(int ,int );
int max4(int ,int ,int ,int );
int main()
{
        int a,b,c,d,max;
        printf("输入4个整数: ");
        scanf("%d %d %d %d",&a,&b,&c,&d);
/***********FOUND***********/
        max4(a, c, b, d);
        printf("最大数是 %d \n",max);
        return 0;
} 
int max4(int a,int b,int c,int d)
{
        int m,n; 
        m=max2(a,b);
        n=max2(c,d);
/***********FOUND***********/
        return max4(m,n);
}
int max2(int a,int b)
{
/***********FOUND***********/
        return a>=b? b: a; 
}