/*-------------------------------------------------------
【程序改错】
---------------------------------------------------------
题目：下列给定程序的功能是：读入一个整数k(2≤k≤10000)，输出它的所有质因
      子(即所有为素数的因子)。
例如：若输入整数2310，则应输出：2,3,5,7,11。
-------------------------------------------------------*/
#include  <conio.h>
#include  <stdio.h>
/***********FOUND***********/
IsPrime(itn n);
{ 
        int i,m;
        m=1;
        for(i=2;i<n;i++)
/***********FOUND***********/
                if!(n%i)
                {
                        m=0;
                        break;
                }
        return(m);
}
main()
{ 
        int j,k;
        printf("\nPlease enter an interger number between 2 and 10000: ");
        scanf("%d",&k);
        printf("\nThe prime factor(s) of %d is(are): ",k);
        for(j=2;j<k;j++)
                 if((!(k%j)) && (IsPrime(j)))
                        printf("%4d,",j);
        printf("\n");
}