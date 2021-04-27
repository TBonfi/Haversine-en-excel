## Formula de distancia en VBA

This is the code you'll need to add a personal formula in Excel VBA to calculate distances in KM between two differents pairs of latitude, longitude coordinates(as in -XX,XXXXXX;-XX,XXXXXX / -XX,XXXXXX;-XX,XXXXXX kind of data) and use the earth radius as an optional input. For having an idea of performance I've used It in an excel spreadsheet to make +5,000,000 calculations in a matrix and It only takes about 30-40 seconds.

It's kind of an amateur work but it serves It's purpose, I work with It on a daily basis and It's good enough for most calculations.

To add this formula you'll need to enable macros an add a module in the file you'll be using(or you can add It on your personal macro book). For more reference just google "How do I add VBA code to my excel" and you'll find lots of tutorials.

Remember that the Radius parameter is optional, which means that you can either use or ignore It. As I only require a 10mts accuracy I default It to the radius that suited me best, but you should either use an explicit value or edit the formula for easiness of use and time saving.

I hope that you'll find It interesting and useful, let me know any comments and look for me in Linkedin and Medium, where I plan to further extend the explanation and uses I give to this utility.

------------------------------------------------------------------------------------------------------------------------------------------

Este es el código que necesitás para tener una fórmula personalizada en Excel VBA para calcular distancia en KM entre dos pares de latitudes y longitudes (En formato -XX,XXXXXX;-XX,XXXXXX / -XX,XXXXXX;-XX,XXXXXX) y usar el radio terrestre como un dato opcional. Para tener una idea de la performance, yo llegué a usarlo en una planilla para calcular +5.000.000 de distancias en una matriz y solo toma unos 30-40 segundos

Es un trabajo rudimentario que cumple su propósito, yo lo uso de forma diaria y es bastante bueno para la mayoría de los cálculos que necesito.

Para agregar la fórmula vas a necesitar tener habilitadas las macros y agregar un módulo en el archivo que vas a usar (o agregarla en tu libro personal de macros). Para más referencias googleá "Como agregar código VBA a mi excel" y vas a encontrar muchos tutoriales.

Recordá que el parámetro "radio" es opcional, así que podés usarlo o ignorarlo y obtener un resultado igualmente. Como yo necesito solo 10mts de precisión lo dejé defaulteado al radio que me queda mejor para Argentina, pero deberías usar un valor explícito o editar el código para más facilidad y ahorro de tiempo.

Ojalá lo encuentren interesante y útil, los comentarios son más que bienvenidos and búsquenme en Linkedin o Medium, donde planeo extenderme en la explicación y contar los usos que le doy a esta utilidad.
