------------------------------------------------------
-- Este ejemplo muestra una lista de todos los generos
-- y del numero de archivos que tienen asociados.
------------------------------------------------------
SELECT	genre.genre AS genero, 
		COUNT(file.id_genre) AS total
FROM	file
RIGHT JOIN genre ON (file.id_genre=genre.id_genre)
GROUP BY genre.genre, file.id_genre
