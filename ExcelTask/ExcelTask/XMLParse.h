#pragma once

#include <MsXml.h>
#include <atlbase.h>
#include <string>

using namespace std;

class XMLParse
{
private:
	string dirname; //адрес папки, где находятся файлы листов
	string outfile; //файл, куда будут записывать результаты
	string filename; //файл xlsx
	string *sheetname; //имена рабочих книг
	string *sheetid;  // id рабочих книг для их дальнейшего сопоставления
	string *relsid;  // id рабочих книг в файле \xl\_rels\workbook.xml.rels
	string *path;     //пути до рабочих книг
	long sheetscount;  //количество рабочих книг

	void DeleteTemp();   //удаление временных файлов, появившихся после разархивирования
	void Initialize();   //инициализация всех значений, связанных с расположением листв, их названием и т.д.

	CComPtr <IXMLDOMDocument> XmlDOMSheet; //DOM-объект для работы с XML-файлом - конкретным листом таблицы
	CComPtr <IXMLDOMDocument> XmlDOMWorkbook; //DOM-объект для работы с XML-файлом, хранящим список листов таблицы и их названия
	CComPtr <IXMLDOMDocument> XmlDOMRels; //DOM-объект для работы с XML-файлом, хранящим адреса листов таблицы \xl\_rels\workbook.xml.rels
	CComPtr <IXMLDOMDocument> XmlDOMString; //DOM-объект для работы с XML-файлом, хранящим строковые значения
public:
	XMLParse(string dir, string file, string out);
	~XMLParse(void);
	void ParseFiles();								//собственно сам парсер
};

