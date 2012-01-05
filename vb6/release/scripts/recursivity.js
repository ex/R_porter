
function accumulate (combiner, nullValue, list)
{
	if (list.length == 0)
		return nullValue;

	var first = list.shift ();

	return combiner (first, accumulate (combiner, nullValue, list));
}

function sumOfSquares (list)
{
	return (accumulate (function(x, y) {return (x*x + y);}, 0, list));
}

function main ()
{
	Console.alert (sumOfSquares([1,2,3,4,5]));
}
