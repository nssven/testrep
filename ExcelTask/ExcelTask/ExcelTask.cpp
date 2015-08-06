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
//!!! � ����� ���������? ��� ������������ ������ � main, ��� � ����������.
string filename;			//��� �������� ����� xlsx
string dirname;				//��� �����, ��� �� ���������
string outfile;				//��� ��������� �����

//������� ���������������� xlsx-�����. �� ����� ��� �����, ������� ����� � ��� ����� (����� ���� ����� � ����� ����� �� �� �������)
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
		cout<<"���������� ������� ����";
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
	if (zr!=ZR_OK) cout<<"������ ��� �������� �����";
}

void main()
{
	//�������� ������� ��������� � ������ ��� ����� xlsx
	setlocale(LC_ALL, "Russian");
	//!!! ������ ������, ���� �� ����� � ���������� �������� � ������.
	filename = new char(); //!!! ��� ���-�� ����� ��������. ������������� string-�� �� ������� �������������. 
	//!!! �� ��� � �� ����. ��� ��� � ���� ���� � ������ �������� � �����.
	cout<<"������� ��� xlsx-�����: ";
	getline(cin,filename);
	cout<<"������� ��� ��������� �����: ";
	getline(cin,outfile);

	int x=filename.length();
	dirname=filename;
	//!!! �������������� _splitpath(). �� ���� ���������� ���������
	for(int i=x-1;i>-1;i--)
		if(dirname[i]!='\\') dirname.erase(i,1);
			else break;
	dirname.erase(dirname.length()-1,2);
	//������������� ����
	UnZipXLSX(filename.c_str(),dirname.c_str()); 	
	//�������� ������ � xml-�������
	XMLParse *parse= new XMLParse(dirname,filename, outfile);
	parse->ParseFiles();
	parse->~XMLParse(); //!!! �������� �������� ����������� ����� delete, � �� ����� ������ �����������
	cout<<"�������� ������� ���������";
	getch();
	return;
}

