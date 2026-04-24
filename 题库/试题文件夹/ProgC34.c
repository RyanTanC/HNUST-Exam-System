#include <stdio.h>
#include <string.h>
void fun(char *tt, int pp[])
{
/**********Program**********/
/**********  End  **********/
}
void main( )
{
char aa[1000] ;
int  bb[26], k ;
printf( "\nPlease enter  a char string:" ) ;
scanf("%s", aa) ;
fun(aa, bb ) ;
for ( k = 0 ; k < 26 ; k++ )
printf ("%d ", bb[k]) ;
printf( "\n" ) ;
}