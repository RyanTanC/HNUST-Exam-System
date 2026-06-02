/*------------------------------------------------
【程序设计】
--------------------------------------------------
功能：求出N×M整型数组的最大元素及其所在的行坐标及
      列坐标（如果最大元素不唯一，选择位置在最前面
      的一个）。
例如：输入的数组为:
                   1   2   3
                   4   15  6
                   12  18  9
                   10  11  2
     求出的最大数为18,行坐标为2，列坐标为1。
------------------------------------------------*/
#include <stdio.h>
#define N 4
#define M 3
int Row,Col;
int fun(int array[N][M])
{
/**********Program**********/
/**********  End  **********/
}
void main()
{
        int a[N][M],i,j,max;
        printf("input a array:");
        for(i=0;i<N;i++)
                for(j=0;j<M;j++)
                        scanf("%d",&a[i][j]);
                for(i=0;i<N;i++)
                {
                        for(j=0;j<M;j++)
                        printf("%d ",a[i][j]);
                        printf("\n");
                }
        max=fun(a);
        printf("max=%d,row=%d,col=%d",max,Row,Col);
}