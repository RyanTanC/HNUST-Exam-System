/*-------------------------------------------------------
【程序设计】
---------------------------------------------------------
题目：主函数main()中一维数组a为测试数据，程序主要判断数组
      a中整数是否是完全数，并将完全数保存在数组b中，然后
      将数组b中最小的完全数移到第一个位置上,其他元素后移，
      其相对位置不变。
完全数：若一个正整数x的所有因子（包括1，不包括x自身）之和
      等于x自身，则称x是一个完全数。例如：28的因子有
      1,2,4,7,14且1+2+4+7+14=28，因此28是一个完全数。
编写程序：
    1. 编写函数int factor(int x[],int n,int y[]),
       在数组x前n个数中查找完全数，将所有完全数保存在数组
       y中，函数返回完全数的个数。
    2. 编写void MoveElement(int y[],int n)，将数组y中n个元
       素最小的完全数移到第一个位置上，其他元素后移后相对
       位置不变。
     测试数据：5,496,23,28,6,8010,2012,1998,1975,2011
     运行结果：Before Moving:
               b[0]=496
               b[1]=28
               b[2]=6
               After Moving:
               b[0]=6
               b[1]=496
               b[2]=28
---------------------------------------------------------
注意：请勿改动主函数main()中的任何语句。
-------------------------------------------------------*/
#include<stdio.h>
#include<stdlib.h>
#include<conio.h>
#include<ctype.h>
#include<string.h>
int factor(int x[],int n,int y[])
{
/**********Program**********/
/**********  End  **********/
}
void MoveElement(int y[],int n)
{
/**********Program**********/
/**********  End  **********/
}
int main()
{
    int a[10]={5,496,23,28,6,8010,2012,1998,1975,2011},b[10]={0},n=10,m,i;
    FILE *fp;
    if((fp=fopen("DATA.TXT","w"))==NULL)
    {
        printf("File open error\n");
        exit(0);
    }
    m=factor(a,n,b);
    printf("Before Moving:\n");  /*移动前*/
    fprintf(fp,"Before Moving:\n");
    for(i=0;i<m;i++)
    {
        printf("b[%d]=%d\n",i,b[i]);
        fprintf(fp,"b[%d]=%d\n",i,b[i]);
    }
    MoveElement(b,m);
    printf("After Moving:\n");  /*移动后*/
    fprintf(fp,"After Moving:\n");
    for(i=0;i<m;i++)
    {
        printf("b[%d]=%d\n",i,b[i]);
        fprintf(fp,"b[%d]=%d\n",i,b[i]);
    }
    fclose(fp);
    getch();
    return 0;
}