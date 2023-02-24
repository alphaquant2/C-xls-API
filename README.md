# ConsoleApplication2This programme demonstrates an API wrapper allowing to write Excel files using an intuitive grammar similar to the stdio formulation

std::cout << variable << std::endl;

2 distinct libraries are used for the demonstration, both are installed through the Vcpkg package manager

    xlsxio
    libxlsxwriter (.vcpkg install libxlsxwriter:x64-windows xlsxio:x64-windows)

Usage: ici on va créer une classe xlsxwriter pour gérer les écritures des fichiers Excel

setcursor(x,y) pour positionner le curseur dans le tableau

write() écrit horizontalement en décalant le curseur vers la droite

writeln() écrit et passe à la ligne suivante vers le bas

l'opérateur << permet à la manière de cout d'écrire horizontalement dans des cellules de gauche à droite

l’opérateur || permet d’écrite un vecteur vecticalement

l'opération close() est appelée automatiquement à la destruction de la variable. Génère une exception si on l'appelle plusieurs fois
This programme demonstrates an API wrapper allowing to write Excel files using an intuitive grammar similar to the stdio formulation

std::cout << variable << std::endl;

2 distinct libraries are used for the demonstration, both are installed through the Vcpkg package manager

    xlsxio
    libxlsxwriter (.vcpkg install libxlsxwriter:x64-windows xlsxio:x64-windows)

Usage: ici on va créer une classe xlsxwriter pour gérer les écritures des fichiers Excel

setcursor(x,y) pour positionner le curseur dans le tableau

write() écrit horizontalement en décalant le curseur vers la droite

writeln() écrit et passe à la ligne suivante vers le bas

l'opérateur << permet à la manière de cout d'écrire horizontalement dans des cellules de gauche à droite

l’opérateur || permet d’écrite un vecteur vecticalement

l'opération close() est appelée automatiquement à la destruction de la variable. Génère une exception si on l'appelle plusieurs fois
