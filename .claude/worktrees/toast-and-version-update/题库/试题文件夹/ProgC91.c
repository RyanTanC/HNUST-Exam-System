/*-------------------------------------------------------
【程序填空】
---------------------------------------------------------
功能：输入一个整数，计算它可能是哪两个整数的平方和，并打印
     结果数据。
     如：34是5和3或3和5的平方和。
-------------------------------------------------------*/
#include  <stdio.h>           /* for i/O functions        */
#include  <stdlib.h>          /* for atoi()               */
#include  <math.h>            /* for sqrt()               */
void  main(void)
{
  int  given;              /* the given number         */
  int  row, column;        /* row and column indicators*/
  int  count;              /* number of solutions      */
  char line[100];
  printf("\nRepresenting a Given Number as the Sum of Two Squares");
  printf("\n=====================================================\n");
  printf("\nAn integer Please ---> ");
  gets(line);
  given = atoi(line);
  printf("\nCount      X      Y");
  printf("\n-----  -----  -----");
  row    = 1;              /* starts from far enough   */
  column = (int) (sqrt((double) given) + 0.5);
  count  = 0;              /* so solution yet          */
  while (row <= given && column > 0)  /* scan down...  */
    if (row*row + column*column == given) 
    {
      /***********SPACE***********/
      【?】;                                        
      printf("\n%5d%7d%7d", count, row, column);
      row++;
      column--;
    }
    else if (row*row + column*column > given)
      /***********SPACE***********/
      【?】;                                        
    else
      /***********SPACE***********/
      【?】;                                        
  if (count == 0)
    printf("\n\nSorry, NO ANSWER found.");
  else
    printf("\n\nThere are %d possible answers.",count);
}