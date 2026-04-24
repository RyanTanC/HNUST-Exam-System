#include <stdio.h>
#include <conio.h>
/**********FOUND**********/
double fun( r)
{
double s;
/**********FOUND**********/
s=1/2*3.14159* r * r;
/**********FOUND**********/
return r;
}
main()
{
float x;
printf ( "Enter x: ");
scanf ( "%f", &x );
printf (" s = %f\n ", fun ( x ) );
}