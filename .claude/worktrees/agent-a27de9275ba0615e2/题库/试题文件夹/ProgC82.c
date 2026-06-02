/*-------------------------------------------------------
【程序设计】
---------------------------------------------------------
题目：主函数main()中a为M行N列二维整型数组，保存M*N个测试
      数据。首先将二维数组中每一行最大元素移动到行尾，最
      大元素和该行其他元素的相对位置保持不变。再对M个行尾
      元素升序排序，同一行元素随行尾元素一起调整，和行尾
      元素保持在同一行中。
例如：下列5*4个数据
        3     2     5     4
        7    21    11     6
        9     4     2     1
        2     8     7     6
       10    12    23    18
      按上述规则每行最大元素移动到行尾后结果为
        4     3     2     5
       11     6     7    21
        4     2     1     9
        7     6     2     8
       18    10    12    23
      按行尾元素升序对二维数组排序后结果为
        4     3     2     5
        7     6     2     8
        4     2     1     9
       11     6     7    21
       18    10    12    23
编写程序：
      1、编写void max_matrix(int a[][N],int m,int n)函数，
         按上述规则移动每一行的最大元素。
      2、编写void sort_matrix(int a[][N],int m,int n)函数，
         按行尾元素升序对二维数组排序。
---------------------------------------------------------
注意：请勿改动主函数main()中的任何语句。
-------------------------------------------------------*/
#include<stdio.h>
#include<stdlib.h>
#include<conio.h>
#include<ctype.h>
#include<math.h>
#define M 5
#define N 4
void max_matrix(int a[][N],int m,int n)
{
/**********Program**********/
/**********  End  **********/
}
void sort_matrix(int a[][N],int m,int n)
{
/**********Program**********/
/**********  End  **********/
}
void main()
{
    int  a[M][N]={{3,2,5,4},{7,21,11,6},{9,4,2,1},{2,8,7,6},{10,12,23,18}};
    int i,j;
    FILE *fp;
    if((fp=fopen("data.txt","w"))==NULL)
    {
        printf("File open error\n");
        exit(0);
    }
    max_matrix(a,M,N);
    fprintf(fp,"After moving:\n");
    printf("After moving:\n");
    for(i=0;i<M;i++)
    {
        for(j=0;j<N;j++)
        {
            fprintf(fp,"%4d",a[i][j]);
            printf("%4d",a[i][j]);
        }
        fprintf(fp,"\n");
        printf("\n");
    }
    fprintf(fp,"\n");
    printf("\n");
    sort_matrix(a,M,N);
    fprintf(fp,"After sorting:\n");
    printf("After sorting:\n");
    for(i=0;i<M;i++)
    {
        for(j=0;j<N;j++)
        {
            fprintf(fp,"%4d",a[i][j]);
            printf("%4d",a[i][j]);
        }
        fprintf(fp,"\n");
        printf("\n");
    }
    fclose(fp);
}