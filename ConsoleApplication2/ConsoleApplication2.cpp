/* This programme demonstrates an API wrapper allowing to write Excel files
* using an intuitive grammar similar to the stdio formulation
* std::cout << variable << std::endl;
*
2 distinc libraries are used for the demonstration, both are installed through the Vcpkg package manager
- xlsxio
- libxlsxwriter
(\.vcpkg install libxlsxwriter:x64-windows xlsxio:x64-windows)

Created by Jérôme Garcia on 2021-01-11
Updated by Jérôme Garcia on 2021-08-04 : intégration de xlsxio par Vcpkg
Updated by Jérôme Garcia on 2023-02-11 : ajout de graphiques dans un fichier Excel

*/
/**************************************************************************
 * (C) Copyright 2021-2023 by Jérôme Garcia. All Rights Reserved.         *
 *                                                                        *
 * DISCLAIMER: The author of this software has used his best efforts in   *
 * preparing the material. These efforts include the	  		          *
 * development, research, and testing of the theories and programs        *
 * to determine their effectiveness. The author makes				      *
 * no warranty of any kind, expressed or implied, with regard to these    *
 * programs or to the documentation contained in these programmes.        *
 * The author shall not be liable in any event for incidental or          *
 * consequential damages in connection with, or arising out of, the       *
 * furnishing, performance, or use of these programs.                     *
 **************************************************************************/


#define _CRT_SECURE_NO_WARNINGS
#include <iostream>
#include <string>
#include <vector>

#include <cstdlib>
#include <Windows.h>
#include <map>
#include <sstream>
#include <filesystem>

using namespace std;
using vecteur = vector<double>;

class image {
	string nomFichier;
public:
	image(const string& str = "") :nomFichier(str) {}
	string nom() const { return nomFichier; }
};

/*UTILISATION DE LA LIBRAIRIE XLSXIO +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++*/
#include "xlsxio_read.h"
#include "xlsxio_write.h"

/*UTILISATION DE LA LIBRAIRIE LIBXLSXWRITER +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
* utilise la librairie xlsxwriter installée avec vcpkg
* cette librairie est très complète et un manuel est disponible sur https ://libxlsxwriter.github.io/index.html
*/
#include "xlsxwriter.h" 

/* ici on va créer une classe xlsxwriter pour gérer les écritures des fichiers Excel
* setcursor(x,y)	pour positionner le curseur dans le tableau
* write() écrit horizontalement en décalant le curseur vers la droite
* writeln() écrit et passe à la ligne suivante vers le bas
* l'opérateur << permet à la manière de cout d'écrire dans des cellules de gauche à droite
* l’opérateur || permet d’écrite un vecteur vecticalement
* l'opération close() est appelée automatiquement à la destruction de la variable. Génère une exception si on l'appelle plusieurs fois
*/

class xlsxwriter {
private:
	lxw_workbook* workbook{ NULL };		// workbook courant
	lxw_worksheet* worksheet{ NULL };	// onglet courant
	lxw_format* format{ NULL };			// le type de format courant
	map<string, lxw_worksheet*> onglet; // dictionnaire créé pour gérer les onglets
	map<string, lxw_format*> mapF;		// dictionnaire créé pour gérer les formats
	int _i{ 0 }, _j{ 0 };
	void close() { workbook_close(workbook); } // pour éviter de l'appeler autrement qu'à la destruction de la variable
	void write(string txt) { worksheet_write_string(worksheet, _j, _i++, txt.c_str(), format); }
	void write(double x) { worksheet_write_number(worksheet, _j, _i++, x, format); }
	void write(int x) { worksheet_write_number(worksheet, _j, _i++, x, format); }
	void write(vecteur V) { for (auto& x : V) write(x); }
	void write(vector<string> V) { for (auto& x : V) write(x); }
	void write(const image& img) { worksheet_insert_image(worksheet, _j, _i++, img.nom().c_str()); }
	void writeln(string txt) { worksheet_write_string(worksheet, _j++, _i, txt.c_str(), format); }
	void writeln(double x) { worksheet_write_number(worksheet, _j++, _i, x, format); }
	void writeln(int x) { worksheet_write_number(worksheet, _j++, _i, x, format); }
	void writeln(vecteur V) { for (auto& x : V) writeln(x); }
	void writeln(vector<string> V) { for (auto& x : V) writeln(x); }
	void writeln(const image& img) {
		if (filesystem::exists(img.nom()))
			worksheet_insert_image(worksheet, _j++, _i, img.nom().c_str());
	}
public:
	xlsxwriter(string nomFichier, string sheet = "Feuille 1") : workbook(workbook_new(nomFichier.c_str())), worksheet(workbook_add_worksheet(workbook, sheet.c_str())), onglet{ { sheet,worksheet } }, mapF{ {"défaut",format} } {}
	void addSheet(string sheet) { worksheet = workbook_add_worksheet(workbook, sheet.c_str()); onglet.insert({ sheet,worksheet }); } // l'onglet courant devient le nouvel onglet
	void addFormat(string form) { format = workbook_add_format(workbook); mapF.insert({ form,format }); }
	void setcursor(int i, int j) { _i = i; _j = j; }

	template <typename T>
	xlsxwriter& operator << (const T& t) { write(t); return *this; }
	xlsxwriter& operator << (std::ostream& (*f)(std::ostream&)) // redéfinition de l'opérateur endl, l'opération lance également le flush du stream
	{
		if (f == (std::basic_ostream<char>&(*)(std::basic_ostream<char>&)) & std::endl)
		{
			_i = 0;// on revient complètement à gauche du tableau
			_j++;
		}
		return *this;
	}
	template <typename T>
	xlsxwriter& operator || (const T& t) { writeln(t); return *this; }

	~xlsxwriter() { close(); } // s'appelle automatiquement à la clôture
};

constexpr int nmax = 1024;// nombre de caractères maximum dans une cellule

vector<vector<string>> ReadFile(string nomFichier, string onglet = ""s) {//utilise la librairie xlsxio
	vector<vector<string>>ligne;
	xlsxioreader xlsxioread;
	if ((xlsxioread = xlsxioread_open(nomFichier.c_str())) == NULL) {
		fprintf(stderr, "Error opening .xlsx file\n");
		return ligne;
	}

	auto pvalue = make_unique<char[]>(nmax);
	auto value = pvalue.get();

	xlsxioreadersheet sheet;

	const char* sheetname = (onglet.length() > 0) ? onglet.c_str() : NULL;
	if (onglet.length() > 0)
		std::cout << "Contenu de l'onglet: " << sheetname << endl;
	else
		std::printf("Contenu du premier onglet:\n");
	if ((sheet = xlsxioread_sheet_open(xlsxioread, sheetname, XLSXIOREAD_SKIP_EMPTY_ROWS)) != NULL) {
		//read all rows
		while (xlsxioread_sheet_next_row(sheet)) {
			//read all columns
			vector<string> item;
			while ((value = xlsxioread_sheet_next_cell(sheet)) != NULL) {
				printf("%s\t", value);
				item.push_back(value);
			}
			printf("\n");
			ligne.push_back(item);
		}
		xlsxioread_sheet_close(sheet);
	}

	xlsxioread_close(xlsxioread);
	return ligne;
}

vector<string> ListerOnglets(string nomFichier) {//utilise la librairie xlsxio
	vector<string> liste;
	xlsxioreader xlsxioread;
	if ((xlsxioread = xlsxioread_open(nomFichier.c_str())) == NULL)
		fprintf(stderr, "Error opening .xlsx file\n");
	else {
		//list available sheets
		xlsxioreadersheetlist sheetlist;
		const char* sheetname = new char[nmax];
		if ((sheetlist = xlsxioread_sheetlist_open(xlsxioread)) != NULL) {
			while ((sheetname = xlsxioread_sheetlist_next(sheetlist)) != NULL)
				liste.push_back(sheetname);
			xlsxioread_sheetlist_close(sheetlist);
		}
		//clean up
		xlsxioread_close(xlsxioread);
		delete[] sheetname;
	}
	return liste;
}

int TestWriteFile(string nomFichier) {// uses xlsxwriter library and smart pointers
	xlsxwriter XW(nomFichier, "démo");
	XW.setcursor(0, 1);
	XW << "hello" << "to everybody" << endl;
	XW << endl;					// line feed
	XW << 3.1415;
	vecteur V{ 1.2,3.5, -6, 7.2,12.22 };
	XW || V;					// writes the vector vertically
	XW.setcursor(0, 5);			// moves to a particular cell
	XW << V;					// writes the vector horizontally
	XW << image("image.jpg");	// inserts a picture at the current position
	return EXIT_SUCCESS;
}

string pwd() //donne le répertoire courant (print working directory)
{
	return filesystem::current_path().string();
}

int main()
{
	/* lecture/écriture de fichiers Excel */
	cout << "current path=" << pwd() << endl;
	std::cout << "Programme de test de lecture de fichiers Excel\n";
	string nomFichier = "liste.xlsx";
	string onglet = "liste";

	/* first test: list the tabs of an Excel file */
	auto VS = ListerOnglets(nomFichier);
	std::cout << "Liste des onglets du fichier " << nomFichier << endl;
	for (auto& x : VS)
		std::cout << x << endl;

	/* Second test: read the contyent of a particular tab*/
	vector<vector<string>>V2 = ReadFile(nomFichier, onglet);// les cellules de l'onglets sont placées dans un tableau double de string

	/* third test: write an Excel file*/
	std::cout << "test d'écriture de fichier. Consulter le résultat dans demo_file.xlsx" << endl;
	TestWriteFile("demo_file.xlsx");

	return EXIT_SUCCESS;
}

