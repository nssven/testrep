#include <windows.h>
#include <tchar.h>
#include <conio.h>
#include <iostream>
#include <string>
#include <cstdio>
#include "zip.h"
#include "unzip.h"
#include "XMLParse.h"


using namespace std;
//!!! А зачем глобально? Они используются только в main, там и определять.
string filename;			//имя текущего файла xlsx
string dirname;				//имя папки, где он находится
string outfile;				//имя выходного файла

//функция разархивирования xlsx-файла. на входе имя файла, которое ввели и имя папки (чтобы было проще и потом снова ее не считать)
void UnZipXLSX(const char *filename, const char *dirname)
{
	HZIP hz;
	HANDLE hf; 
	DWORD writ; 
	ZRESULT zr; 
	ZIPENTRY ze; 
	TCHAR m[1024];
	hz=OpenZip(filename,0);
	if (hz==0) 
	{
		cout<<"Невозможно открыть файл";
		return;
	}
    zr=SetUnzipBaseDir(hz,dirname); 
    zr=GetZipItem(hz,-1,&ze); 
    int numitems=ze.index; 
	for (int i=0; i<numitems; i++)
    { 
		zr=GetZipItem(hz,i,&ze); 
		zr=UnzipItem(hz,i,ze.name);
    }
    zr=CloseZip(hz); 
	if (zr!=ZR_OK) cout<<"Ошибка при закрытии файла";
}

void main()
{
	//поставим русскую кодировку и введем имя файла xlsx
	setlocale(LC_ALL, "Russian");
	//!!! Кстати говоря, путь до файла с кириллицей приводит к ошибке.
	filename = new char(); //!!! это что-то очень странное. Инициализация string-ов по другому производиться. 
	//!!! Да это и не надо. Тут еще и байт один в памяти теряется в итоге.
	cout<<"Введите имя xlsx-файла: ";
	getline(cin,filename);
	cout<<"Введите имя выходного файла: ";
	getline(cin,outfile);

	int x=filename.length();
	dirname=filename;
	//!!! воспользуйтесь _splitpath(). Не надо изобретать велосипед
	for(int i=x-1;i>-1;i--)
		if(dirname[i]!='\\') dirname.erase(i,1);
			else break;
	dirname.erase(dirname.length()-1,2);
	//разархивируем файл
	UnZipXLSX(filename.c_str(),dirname.c_str()); 	
	//начинаем работу с xml-файлами
	XMLParse *parse= new XMLParse(dirname,filename, outfile);
	parse->ParseFiles();
	parse->~XMLParse(); //!!! удаление объектов выполняется через delete, а не путем вызова деструктора
	cout<<"Операция успешно выполнена";
	getch();
	return;
}

