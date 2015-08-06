#pragma once

#include <MsXml.h>
#include <atlbase.h>
#include <string>

using namespace std;

class XMLParse
{
private:
	string dirname; //����� �����, ��� ��������� ����� ������
	string outfile; //����, ���� ����� ���������� ����������
	string filename; //���� xlsx
	string *sheetname; //����� ������� ����
	string *sheetid;  // id ������� ���� ��� �� ����������� �������������
	string *relsid;  // id ������� ���� � ����� \xl\_rels\workbook.xml.rels
	string *path;     //���� �� ������� ����
	long sheetscount;  //���������� ������� ����

	void DeleteTemp();   //�������� ��������� ������, ����������� ����� ����������������
	void Initialize();   //������������� ���� ��������, ��������� � ������������� �����, �� ��������� � �.�.

	CComPtr <IXMLDOMDocument> XmlDOMSheet; //DOM-������ ��� ������ � XML-������ - ���������� ������ �������
	CComPtr <IXMLDOMDocument> XmlDOMWorkbook; //DOM-������ ��� ������ � XML-������, �������� ������ ������ ������� � �� ��������
	CComPtr <IXMLDOMDocument> XmlDOMRels; //DOM-������ ��� ������ � XML-������, �������� ������ ������ ������� \xl\_rels\workbook.xml.rels
	CComPtr <IXMLDOMDocument> XmlDOMString; //DOM-������ ��� ������ � XML-������, �������� ��������� ��������
public:
	XMLParse(string dir, string file, string out);
	~XMLParse(void);
	void ParseFiles();								//���������� ��� ������
};

