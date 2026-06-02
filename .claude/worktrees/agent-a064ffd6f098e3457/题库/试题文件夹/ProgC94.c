/*-------------------------------------------------------
【程序设计】
---------------------------------------------------------
题目：主函数main()中字符串a，b为测试数据，字符串a中字符已
      按ASCII码值降序排列，将字符串b中ASCII码为偶数的字符
      依次插入到字符串a中，使字符串a中字符仍按ASCII码值降
      序排列。再删除字符串a中的重复字符，使相同字符只出现
      一次。
编写程序：
     1、编写void insert(char a[],char b[])函数，按上述规
        则将字符串b中字符插入到字符串a中。
     2、编写 void del(char a[])函数，删除字符串a中的重复
        字符，使相同字符只出现一次。
---------------------------------------------------------
注意：请勿改动主函数main()中的任何语句。
-------------------------------------------------------*/
#include<stdio.h>
#include<stdlib.h>
#include<conio.h>
#include<ctype.h>
#include<math.h>
void insert(char a[],char b[])
{
/**********Program**********/
/**********  End  **********/
}
void del(char a[])
{
/**********Program**********/
/**********  End  **********/
}
void main(){
    char a[30]="9777532";
    char b[10]= "987645231";
    FILE *fp;
    if((fp=fopen("data.txt","w"))==NULL)
    {
        printf("File open error\n");
        exit(0);
    }
    insert(a,b);
    fprintf(fp,"After insterting:");
    printf("After insterting:");
    fprintf(fp,"%s\n",a);
    printf("%s\n",a);
    del(a);
    fprintf(fp,"After deleting:");
    printf("After deleting:");
    fprintf(fp,"%s\n",a);
    printf("%s\n",a);
    fclose(fp);
}