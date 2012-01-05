
// Continuations
// http://www.innoq.com/blog/st/2005/04/13/continuations_for_curmudgeons.html 
function makeFibonacci()
{
	var i=0;
    var j=1;
	fib = function()
	{
		var k = i;
        i = j;
    	j = k + j;
    	return k;
    }
	return fib;
}

function fibContinuations( fibonacci )
{
	var k = fibonacci();
	while( k < 1000 )
	{
		Console.echo( k );
        k = fibonacci();
	}
}

function fibBinet( fibonacci )
{
	var k = 0;
	var fib = 0;
	while( fib < 1000 )
	{
		Console.echo( fib );
		fib = fibonacci(k);
        k++;
	}
}

// Return the n-th Fibonacci number using Binet's theorem
var C1	= Math.pow( 5.0, 0.5 );
var PHI	= ( 1 + C1 )/2;

function binet(n)
{
  value = ( ( Math.pow( PHI, n + 1 ) - 
			  Math.pow( ( 1 - PHI ), n + 1 ) ) )/C1;
  return parseInt( value );
} 

function main()
{
	Console.clrscr();
	fibContinuations( makeFibonacci() );
	Console.echo( "-------------------" );
	fibBinet( binet );
}
