#include "XMLParse.h"
#include <MsXml6.h>
#include <atlbase.h>
#include <iostream>
#include <Windows.h>
#include <string>
#include <cstdlib>
#include <wchar.h>
#include <tchar.h>
#include <fstream>
#include <locale>


using namespace std;

XMLParse::XMLParse(string dir, string file, string out)
{
	//инициализируем DOM-объект для работы с XML-файлом
	CoInitialize(NULL);
	dirname=dir;
	outfile=out;
	filename=file;
	sheetscount=0;
}

XMLParse::~XMLParse(void)
{
}

void XMLParse::Initialize()
{
	HRESULT HR = XmlDOMSheet.CoCreateInstance(__uuidof(DOMDocument));
	XmlDOMString.CoCreateInstance(__uuidof(DOMDocument));
	XmlDOMWorkbook.CoCreateInstance(__uuidof(DOMDocument));
	VARIANT_BOOL bSuccess = false;
	HR = XmlDOMWorkbook->load(CComVariant((string(dirname)+"\\xl\\workbook.xml").c_str()), &bSuccess);
	try
	{
		if ( FAILED(HR) || !bSuccess )
				throw "Невозможно загрузить XML-документ";
	}
	catch(...)
	{
		cout<<"Ошибка при загрузке информации о книге";
		return;
	}
	IXMLDOMNodeList *pNodesRow = NULL;
	IXMLDOMNode *pNodeCell = NULL;
	IXMLDOMNode *pNodeCellName = NULL;
	//находим элементы с тегом <sheets>
	HR=XmlDOMWorkbook->getElementsByTagName(SysAllocString(L"sheet"),&pNodesRow);
	SUCCEEDED(HR) ? 0 : throw HR;
	//находим количество этих элементов
	HR = pNodesRow->get_length(&sheetscount);
	sheetname=new string[sheetscount];
	sheetid=new string[sheetscount];
	path= new string[sheetscount];
	relsid= new string[sheetscount];
	if(SUCCEEDED(HR))
	{	
		pNodesRow->reset();
		for(int i = 0; i < sheetscount; i++)
		{
			//выделяем конкретный элемент
			pNodesRow->get_item(i, &pNodeCell);
			if(pNodeCell)
			{
				VARIANT stname, stid;
				long p;
				IXMLDOMNamedNodeMap *atr;
				HR=pNodeCell->get_attributes(&atr);
				atr->get_length(&p);
				HR = atr->getNamedItem(L"name", &pNodeCellName);
				pNodeCellName->get_nodeValue(&stname);
				HR = atr->getNamedItem(L"r:id", &pNodeCellName);
				pNodeCellName->get_nodeValue(&stid);	
				//записываем в файл содержимое ячейки
				USES_CONVERSION;
					std::string sname(W2A(stname.bstrVal));
					std::string sid(W2A(stid.bstrVal));
				sheetname[i]=sname;
				sheetid[i]=sid;
				pNodeCell->Release();
				pNodeCell = NULL;
			}
		}
	}

	//обнуляем
	pNodesRow->Release();
	pNodesRow = NULL;

	XmlDOMRels.CoCreateInstance(__uuidof(DOMDocument));
	bSuccess = false;
	HR = XmlDOMRels->load(CComVariant((string(dirname)+"\\xl\\_rels\\workbook.xml.rels").c_str()), &bSuccess);
	try
	{
		if ( FAILED(HR) || !bSuccess )
				throw "Невозможно загрузить XML-документ";
	}
	catch(...)
	{
		cout<<"Ошибка при загрузке информации о листах книги";
		return;
	}
	HR=XmlDOMRels->getElementsByTagName(SysAllocString(L"Relationship"),&pNodesRow);
	SUCCEEDED(HR) ? 0 : throw HR;
	long count;
	//находим количество этих элементов
	HR = pNodesRow->get_length(&count);
	if(SUCCEEDED(HR))
	{	
		pNodesRow->reset();
		for(int i = 0; i < sheetscount; i++)
		{
			//выделяем конкретный элемент
			pNodesRow->get_item(i, &pNodeCell);
			if(pNodeCell)
			{
				VARIANT stpath, stid;
				long p;
				IXMLDOMNamedNodeMap *atr;
				HR=pNodeCell->get_attributes(&atr);
				atr->get_length(&p);
				HR = atr->getNamedItem(L"Id", &pNodeCellName);
				pNodeCellName->get_nodeValue(&stid);
				HR = atr->getNamedItem(L"Target", &pNodeCellName);
				pNodeCellName->get_nodeValue(&stpath);			
				USES_CONVERSION;
					std::string sid(W2A(stid.bstrVal));
					std::string spath(W2A(stpath.bstrVal));
					relsid[i]=sid;
					for(int j=0;j<sheetscount;j++)
					{
						if(sheetid[j]==relsid[i])
							path[i]=spath;
					}
				pNodeCell->Release();
				pNodeCell = NULL;
			}
		}
	}
	//обнуляем
	pNodesRow->Release();
	pNodesRow = NULL;

	//открываем текстовый файл для вывода и записываем туда имя файла
	FILE * myfile = fopen(outfile.c_str(), "w");
	string header=filename;
	header="Input: "+header+"\n";
	fputs(header.c_str(),myfile);
	fclose(myfile);									//закрываем файл для записи
}

void XMLParse::DeleteTemp()
{
	//!!! Зря распаковываете прямо рядом с заданием. 
	//!!! Можно было бы создать папку, в которую все сложить, а потом целиком удалить. (ну это как совет)

	//имена файлов и папок для дальнейшего удаления
	string file_n=dirname;
	string folders[3];
	folders[0]=dirname+"\\xl"+ _T('\0');
	folders[1]=dirname+"\\docProps\0"+ _T('\0');
	folders[2]=dirname+"\\_rels\0"+ _T('\0');
	string dir=dirname;
	file_n+="\\[Content_Types].xml";				//удаляем временный файл
	DeleteFile(file_n.c_str());	
	SHFILEOPSTRUCT fo;								//удаляем папки с временными файлами
	ZeroMemory(&fo, sizeof(fo));
	fo.wFunc  = FO_DELETE;
	fo.fFlags = FOF_NOCONFIRMATION | FOF_SILENT;
	fo.hNameMappings = 0;
	fo.lpszProgressTitle = NULL;
	for(int i=0;i<3;i++)
	{
		fo.pFrom  = folders[i].c_str();
		SHFileOperation(&fo);
	}
}


void XMLParse::ParseFiles()
{
	Initialize();
	int ind=0;
	VARIANT_BOOL bSuccess = false;
	FILE * myfile = fopen(outfile.c_str(), "a");
	//рассматриваем каждый лист книги
	for(int i=0;i<sheetscount;i++)
	{
		int k=-1;  //k - индекс элемента, который хранит путь до файла
		//здесь ищет путь до него среди распакованных файлов архива
		for(int j=0;j<sheetscount;j++)
		{	
			if(sheetid[i]==relsid[j])
			{
				k=j;
				break;
			}
		}
		if(k==-1) throw "Произошла неожиданная ошибка"; //в приницпе такого быть не может, если файл корректно записан		
		//загружаем текущий лист
		HRESULT HR = XmlDOMSheet->load(CComVariant((dirname+"\\xl\\"+path[k]).c_str()), &bSuccess);
		try
		{
			if ( FAILED(HR) || !bSuccess )
			throw "Невозможно загрузить XML-документ";
		}
		catch(...)
		{
			cout<<"Ошибка при загрузке листа книги\r\n";
			return;
		}

		//собственно сам парсинг листа
		BSTR bstrItemText, my;
		long value;
		IXMLDOMNodeList *pNodesRow = NULL;
		IXMLDOMNodeList *pNodesCellsName = NULL;
		IXMLDOMNode *pNodeCell = NULL;
		IXMLDOMNode *pNodeCellName = NULL;

		//находим элементы с тегом <row>
		HR=XmlDOMSheet->getElementsByTagName(SysAllocString(L"row"),&pNodesRow);
		SUCCEEDED(HR) ? 0 : throw HR;
		//находим количество этих элементов
		HR = pNodesRow->get_length(&value);
		if(SUCCEEDED(HR))
		{	
			fputs(("-Sheet: "+sheetname[i]+"\n").c_str(), myfile);

			pNodesRow->reset();
			for(int i = 0; i < value; i++)
			{
				//выделяем конкретный элемент
				pNodesRow->get_item(i, &pNodeCell);
				if(pNodeCell)
				{
					//записываем содержимое данной ячейки
					HR = pNodeCell->get_text(&bstrItemText);
					//находим ее дочерние узлы
					pNodeCell->get_childNodes(&pNodesCellsName);
					//вытаскиваем имя данной ячейки
					long t,p;
					VARIANT varValue,text;
					IXMLDOMNamedNodeMap *atr;
					pNodesCellsName->get_length(&t);
					pNodesCellsName->get_item(0,&pNodeCellName);
					HR=pNodeCellName->get_attributes(&atr);
					atr->get_length(&p);
					HR = atr->getNamedItem(L"r", &pNodeCellName);
					pNodeCellName->get_nodeValue(&varValue);

					HR = atr->getNamedItem(L"t", &pNodeCellName);
					if(HR==S_FALSE)
					{
						//записываем в файл содержимое ячейки
						USES_CONVERSION;
						std::string svalue(W2A(bstrItemText));
						std::string scell(W2A(varValue.bstrVal));
						
						fputs(("--"+scell+": "+svalue+"\n").c_str(), myfile);
					}
					else
					{			
						//в ячейке строка, поэтому ищем сапоставление числового значения строке
						bSuccess = false;				
						HR = XmlDOMString->load(CComVariant((dirname+"\\xl\\sharedStrings.xml").c_str()), &bSuccess);
						try
						{
							if ( FAILED(HR) || !bSuccess )
							throw "Невозможно загрузить XML-документ";
						}
						catch(...)
						{
							cout<<"Ошибка при загрузке строковых значений";
							return;
						}

						
						BSTR bstrString;
						long valueS;
						IXMLDOMNodeList *pNodesRowS = NULL;
						IXMLDOMNode *pNodeCellS = NULL;
						VARIANT varValueS; //!!! это что-то лишнее

						//находим элементы с тегом <t>
						HR=XmlDOMString->getElementsByTagName(SysAllocString(L"t"),&pNodesRowS);
						SUCCEEDED(HR) ? 0 : throw HR;
						//находим количество этих элементов
						HR = pNodesRowS->get_length(&valueS);
						
						pNodesRowS->get_item(ind, &pNodeCellS);
						ind++;
						HR = pNodeCellS->get_text(&bstrString);


						//записываем в файл содержимое ячейки
						USES_CONVERSION;
					//	std::wstring svalue(bstrString,SysStringLen(bstrString));

						//!!! так не получиться вывести любой юникод. Надо стремиться наоборот перегонять 
						//!!! все в Юникод (string -> wstring, char -> wchar_t)
						std::string svalue(W2A(bstrString));
						std::string scell(W2A(varValue.bstrVal));

						//попытки записать Unicode-строку
				//		int length = SysStringLen(bstrString); // для размера BSTR
				//		wchar_t *myString = new wchar_t[length+1]; // Примечание: не SysStringLen
				//		wcscpy(myString, bstrString);
				//		fputws(myString,myfile);
						fputs(("--"+scell+": "+svalue+"\n").c_str(), myfile);   //теряются некоторые символы после преобразования BSTR->string			
					}
				}
			}
		}
	}
	fclose(myfile);
	DeleteTemp();
}

