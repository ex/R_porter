//--------------------------------------------------
// This script shows the acumulator
// MORE INFO: http://www.paulgraham.com/icad.html
//--------------------------------------------------

function foo(n) 
{
	// notice we are returning (n + i)
	// but at the same time we are INCREMENTING (n)
	return function (i) { return n += i; }
}

function main ()
{
	f = foo(10);
	Console.clrscr ();
	Console.echo (f(5));	// --> 15
	Console.echo (f(10));	// --> 25 (NOT 20)
}
