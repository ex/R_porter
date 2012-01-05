//-----------------------------------------------------
// Map-Reduce example (Joel Spolsky)
//-----------------------------------------------------
function map( fn, a )
{
	for( k = 0; k < a.length; k++ )
	{
		val = fn( a[k] );
		if( val )	{ a[k] = val; }
	}
}

function reduce( fn, a, init )
{
	var s = init;

	for( k = 0; k < a.length; k++ )
	{
		s = fn( s, a[k] );
	}
	return s;
}

function main()
{
	var a = [1, 2, 3];

	map( function( x ){ return x*2; }, a );
	map( function( x ){ alert( x ); }, a );

	alert( reduce( function( x, y ){ return x + y; }, a, 0 ));
	alert( reduce( function( x, y ){ return x + y; }, a, "" ));

}

