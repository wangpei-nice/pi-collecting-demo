// pisdk-demo.cpp : This file contains the 'main' function. Program execution begins and ends there.
//
//#pragma warning(suppress : 4996)
#include <iostream>
#include <fstream>
#include <sstream>
#include <vector>
#include <chrono>
#include <thread>
//#include <windows.h>
#include <atlbase.h>

#import "pisdkcommon.dll" no_namespace
#import "piTimeServer.dll" no_namespace
#import "pisdk.dll" no_namespace 

using namespace std;

// ��ȡCSV�ļ�
vector<string> readCSV(const string& filename)
{
    vector<string> data;
    ifstream file(filename);
    if (!file)
    {
        cout << "�޷����ļ���" << filename << endl;
        return data;
    }

    string line;
    while (getline(file, line))
    {
        stringstream ss(line);
        string cell;

        if (getline(ss, cell, ','))
        {
            data.push_back(cell);
        }  
    }

    file.close();
    return data;
}

// д��CSV�ļ�
void writeCSV(const string& filename, const vector<vector<string>>& data)
{
    ofstream file(filename, std::ios::app);
    if (!file)
    {
        cout << "�޷������ļ���" << filename << endl;
        return;
    }
    cout << "�����ļ���" << filename << endl;

    for (const auto& row : data)
    {
        for (const auto& cell : row)
        {
            file << cell << ",";
        }
        file << endl;
    }

    file.close();
}


void GetTagsValueByPIPointsTh(PIPointsPtr pPoints, std::vector<string> tags) {
    CoInitialize(NULL);
    try {
        auto start = std::chrono::high_resolution_clock::now();      

        for (int i = 0; i < tags.size(); i++) {
            PIPointPtr pPt = pPoints->GetItem(_bstr_t(tags[i].c_str()));
            _PIDataPtr pData = pPt->GetData();
            _PIValuePtr pValue = pData->GetSnapshot();
            _variant_t var = pValue->GetValue();
        }

        auto end = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);

        std::cout << "����ִ��ʱ��: " << duration.count() << " ����\n";
    }

    catch (_com_error Err)
    {
        _tprintf(_T("Error:%s : 0x%x \n"), (TCHAR*)Err.Description(), Err.Error());
    }

    catch (...)
    {
        std::cout << "Unknown exception caught" << std::endl;
    }
    CoUninitialize();
}


void GetTagsValueTh(std::vector<string> tags) {
    CoInitialize(NULL);
    try {

        IPISDKPtr pSDK(__uuidof(PISDK));

        ServersPtr pServs = pSDK->GetServers();

        ServerPtr pServ = pServs->GetItem("piserver");

        PIPointsPtr pPoints = pServ->PIPoints;

        auto start = std::chrono::high_resolution_clock::now();

        for (int i = 0; i < tags.size(); i++) {
            PIPointPtr pPt = pPoints->GetItem(_bstr_t(tags[i].c_str()));
            _PIDataPtr pData = pPt->GetData();
            _PIValuePtr pValue = pData->GetSnapshot();
            _variant_t var = pValue->GetValue();
        }

        auto end = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);

        std::cout << "����ִ��ʱ��: " << duration.count() << " ����\n";
    }

    catch (_com_error Err)
    {
        _tprintf(_T("Error:%s : 0x%x \n"), (TCHAR*)Err.Description(), Err.Error());
    }

    catch (...)
    {
        std::cout << "Unknown exception caught" << std::endl;
    }
    CoUninitialize();
}

bool ParseCommandLine(LPTSTR szCmdLine, _bstr_t& bServerName,
    _bstr_t& bUserName, _bstr_t& bPassword)
{
    // skip over program directory
    LPTSTR szPtr = _tcstok(szCmdLine, _T(" \n\0"));
    if (!szPtr)
        return false;
    szPtr = _tcstok(NULL, _T(" \n\0"));
    if (!szPtr)
        return false;
    bServerName = szPtr;
    szPtr = _tcstok(NULL, _T(" \n\0"));
    if (!szPtr)
        return false;
    bUserName = szPtr;
    szPtr = _tcstok(NULL, _T("\n\0"));
    if (!szPtr)
        return false;
    bPassword = szPtr;
    return true;
}


// Method 1
int method1()
{
    LPTSTR szCmd = GetCommandLine();
    _bstr_t bServerName, bUserName, bPassword;
    if (!ParseCommandLine(szCmd, bServerName, bUserName, bPassword))
    {
        _tprintf(_T("Usage: AddUser Server User Password\n"));
        return 0;
    }

    CoInitialize(NULL);
    try
    {
        IPISDKPtr pSDK(__uuidof(PISDK));

        ServersPtr pServs = pSDK->GetServers();

        ServerPtr pServ = pServs->GetItem(bServerName);

        _bstr_t connectString("UID=" + bUserName + ";PWD=" + bPassword);



        HRESULT result = pServ->Open(connectString);
        if (result == S_OK) {
            std::cout << "Connect successfully" << std::endl;
        }


        // ����sql��ѯ�ĺ�ʱ
        auto start0 = std::chrono::high_resolution_clock::now();

        _PointListPtr pPointList = pServ->GetPoints((_bstr_t)"tag = 'XTEQ*'", 0); // ��ȡ��λ��ΪXTEQ��ͷ�����е�λ�����64000��

        auto end0 = std::chrono::high_resolution_clock::now();
        auto duration0 = std::chrono::duration_cast<std::chrono::milliseconds>(end0 - start0);

        std::cout << "SQL��ѯʱ��: " << duration0.count() << " ����\n";
        std::cout << "��ѯ�����ݹ���" << pPointList->GetCount() << " ��\n";



        // �����ȡ��λֵ�ĺ�ʱ
        auto start = std::chrono::high_resolution_clock::now();

        _variant_t var;
        VARIANT index;
        VariantInit(&index);
        index.vt = VT_INT;
        vector<vector<string>> datas;
        for (int i = 1; i <= pPointList->GetCount(); ++i) { //ѭ����ȡÿ����λ��ֵ
            index.intVal = i;
            PIPointPtr pPt = pPointList->GetItem(&index);
            var = pPt->Data->GetSnapshot()->Value;
        }
        VariantClear(&index);
        auto end = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);

        std::cout << "����ִ��ʱ��: " << duration.count() << " ����\n";
    }
    catch (_com_error Err)
    {
        _tprintf(_T("Error:%s : 0x%x \n"), (TCHAR*)Err.Description(), Err.Error());
    }
    CoUninitialize();
    return 0;
}


// method 2
int method2()
{
    LPTSTR szCmd = GetCommandLine();
    _bstr_t bServerName, bUserName, bPassword;
    if (!ParseCommandLine(szCmd, bServerName, bUserName, bPassword))
    {
        _tprintf(_T("Usage: AddUser Server User Password\n"));
        return 0;
    }

    CoInitialize(NULL);
    try
    {
        IPISDKPtr pSDK(__uuidof(PISDK));

        ServersPtr pServs = pSDK->GetServers();

        ServerPtr pServ = pServs->GetItem(bServerName);

        _bstr_t connectString("UID=" + bUserName + ";PWD=" + bPassword);



        HRESULT result = pServ->Open(connectString);
        if (result == S_OK) {
            std::cout << "Connect successfully" << std::endl;
        }

        std::vector<string> tags1;
        tags1.reserve(10000);
        // ��test1.csv�ж�ȡ��λ���ƣ��浽һ���б��У�1�����λ
        tags1 = readCSV("test1.csv");
        
        std::vector<string> tags2;
        tags2.reserve(10000);
        // ��test2.csv�ж�ȡ��λ���ƣ��浽һ���б��У�1�����λ
        tags2 = readCSV("test2.csv");


       PIPointsPtr pPoints = pServ->PIPoints;
        // ���������̣߳�����һ��PIPointsPtrȥ��ȡ��λ���ݣ�ÿ���̷ֱ߳��ȡ10000����λ
        std::thread t1(GetTagsValueByPIPointsTh, pPoints, tags1);
        std::thread t2(GetTagsValueByPIPointsTh, pPoints, tags2);

        // �ȴ��߳̽���
        t1.join();
        t2.join();

        cout << "Threads finished." << endl;
    }
    catch (_com_error Err)
    {
        _tprintf(_T("Error:%s : 0x%x \n"), (TCHAR*)Err.Description(), Err.Error());
    }
    CoUninitialize();
    return 0;
}


// method 3
int method3()
{
    CoInitialize(NULL);
    try
    {
        std::vector<string> tags1;
        tags1.reserve(10000);
        // ��test1.csv�ж�ȡ��λ���ƣ��浽һ���б��У�1�����λ
        tags1 = readCSV("test1.csv");

        std::vector<string> tags2;
        tags2.reserve(10000);
        // ��test2.csv�ж�ȡ��λ���ƣ��浽һ���б��У�1�����λ
        tags2 = readCSV("test2.csv");

        // ���������̣߳�ÿ���߳��и��Ե�PISDKʵ��ȥ��ȡ��λ��ÿ���̷ֱ߳��ȡ10000����λ
        std::thread t1(GetTagsValueTh, tags1);
        std::thread t2(GetTagsValueTh, tags2);

        // �ȴ��߳̽���
        t1.join();
        t2.join();

        cout << "Threads finished." << endl;
    }
    catch (_com_error Err)
    {
        _tprintf(_T("Error:%s : 0x%x \n"), (TCHAR*)Err.Description(), Err.Error());
    }
    CoUninitialize();
    return 0;
}


// method 4
int method4()
{
    LPTSTR szCmd = GetCommandLine();
    _bstr_t bServerName, bUserName, bPassword;
    if (!ParseCommandLine(szCmd, bServerName, bUserName, bPassword))
    {
        _tprintf(_T("Usage: AddUser Server User Password\n"));
        return 0;
    }

    CoInitialize(NULL);
    try
    {
        IPISDKPtr pSDK(__uuidof(PISDK));

        ServersPtr pServs = pSDK->GetServers();

        ServerPtr pServ = pServs->GetItem(bServerName);

        _bstr_t connectString("UID=" + bUserName + ";PWD=" + bPassword);



        HRESULT result = pServ->Open(connectString);
        if (result == S_OK) {
            std::cout << "Connect successfully" << std::endl;
        }

        std::vector<string> tags;
        tags.reserve(10000);
        // ��test1.csv�ж�ȡ��λ���ƣ��浽һ���б��У�1�����λ
        tags = readCSV("test1.csv");


        auto start = std::chrono::high_resolution_clock::now();

        PIPointsPtr pPoints = pServ->PIPoints;

        _variant_t var;
        for (int i = 0; i < tags.size(); i++) {
            PIPointPtr pPt = pPoints->GetItem(_bstr_t(tags[i].c_str()));
            _PIDataPtr pData = pPt->GetData(); 
            _PIValuePtr pValue = pData->GetSnapshot(); 
            var = pValue->GetValue();
        }

        auto end = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);

        std::cout << "����ִ��ʱ��: " << duration.count() << " ����\n";
    }
    catch (_com_error Err)
    {
        _tprintf(_T("Error:%s : 0x%x \n"), (TCHAR*)Err.Description(), Err.Error());
    }
    CoUninitialize();
    return 0;
}


int method5() {
    LPTSTR szCmd = GetCommandLine();
    _bstr_t bServerName, bUserName, bPassword;
    if (!ParseCommandLine(szCmd, bServerName, bUserName, bPassword))
    {
        _tprintf(_T("Usage: AddUser Server User Password\n"));
        return 0;
    }

    CoInitialize(NULL);
    try
    {
        IPISDKPtr pSDK(__uuidof(PISDK));

        ServersPtr pServs = pSDK->GetServers();

        ServerPtr pServ = pServs->GetItem(bServerName);

        _bstr_t connectString("UID=" + bUserName + ";PWD=" + bPassword);



        HRESULT result = pServ->Open(connectString);
        if (result == S_OK) {
            std::cout << "Connect successfully" << std::endl;
        }

        PIPointsPtr pPoints = pServ->PIPoints;

        _variant_t var;

        PIPointPtr pPt = pPoints->GetItem(_bstr_t("XTEQ_LJ:1150UZSO112"));
        _PIDataPtr pData = pPt->GetData();
        _PIValuePtr pValue = pData->GetSnapshot();
        var = pValue->GetValue();
        _bstr_t bstrValue(var);
        std::string stringValue(static_cast<const char*>(bstrValue));
        std::cout << "var:" << bstrValue << std::endl;
        pServ->Close();

    }
    catch (_com_error Err)
    {
        _tprintf(_T("Error:%s : 0x%x \n"), (TCHAR*)Err.Description(), Err.Error());
    }
    CoUninitialize();
    return 0;
}


int method6()
{
    LPTSTR szCmd = GetCommandLine();
    _bstr_t bServerName, bUserName, bPassword;
    if (!ParseCommandLine(szCmd, bServerName, bUserName, bPassword))
    {
        _tprintf(_T("Usage: AddUser Server User Password\n"));
        return 0;
    }

    CoInitialize(NULL);
    try
    {
        IPISDKPtr pSDK(__uuidof(PISDK));

        ServersPtr pServs = pSDK->GetServers();

        ServerPtr pServ = pServs->GetItem(bServerName);

        _bstr_t connectString("UID=" + bUserName + ";PWD=" + bPassword);
        HRESULT result = pServ->Open(connectString);
        if (result == S_OK) {
            std::cout << "Connect successfully" << std::endl;
        }

        std::vector<string> tags;
        tags.reserve(10000);
        // ��test.csv�ж�ȡ��λ���ƣ��浽һ���б��У�1�����λ
        tags = readCSV("test.csv");

        cout << "tags.size: " << tags.size() << endl;
        if (tags.empty())
            return -1;

        // ����sql��ѯ�ĺ�ʱ
        auto start0 = std::chrono::high_resolution_clock::now();

        std::string whereCause = "tag = '" + tags[0] + "'";
        _PointListPtr pFinalPointList  = pServ->GetPoints((_bstr_t)whereCause.c_str(), 0);
        PIPointsPtr pPoints = pServ->PIPoints;
        for (int i = 1; i < tags.size(); ++i) {          
            PIPointPtr pPt = pPoints->GetItem(_bstr_t(tags[i].c_str()));
            pFinalPointList->Add(pPt);
        }
   
        auto end0 = std::chrono::high_resolution_clock::now();
        auto duration0 = std::chrono::duration_cast<std::chrono::milliseconds>(end0 - start0);

        std::cout << "SQL��ѯʱ��: " << duration0.count() << " ����\n";
        std::cout << "��ѯ�����ݹ���" << pFinalPointList->GetCount() << " ��\n";


        // �����ȡ��λֵ�ĺ�ʱ
        auto start = std::chrono::high_resolution_clock::now();
    
        struct _NamedValues* errors = nullptr;
        _ListDataPtr pDataList = pFinalPointList->GetData();
        PointValuesPtr pPointValues = pDataList->GetSnapshot(&errors);
        IEnumVARIANTPtr pEnumPoints = pPointValues->_NewEnum;
        

        HRESULT res;
        VARIANT varItem;
        ULONG fetched;
        std::vector<std::vector<string>> data;

        VariantInit(&varItem);
        res = pEnumPoints->Next(1, &varItem, &fetched);
        while (SUCCEEDED(res) && fetched > 0)
        {
            std::vector<string> cells;

            PointValuePtr pPointValue = varItem.punkVal;
            _PIValuePtr pPIValue = pPointValue->GetPIValue();
                               
            _variant_t var;
            var = pPIValue->GetValue();
            _bstr_t bstrValue(var);
            std::string stringValue(static_cast<const char*>(bstrValue));

            PIPointPtr pPIPoint = pPointValue->GetPIPoint();
            
            //std::cout << "name:" << pPIPoint->GetName() << ", var:" << bstrValue << std::endl;
            cells.push_back(string(pPIPoint->GetName()));
            cells.push_back(stringValue);
            data.push_back(cells);

            VariantClear(&varItem);
            VariantInit(&varItem);
            res = pEnumPoints->Next(1, &varItem, &fetched);
        }

        auto end = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);

        std::cout << "����ִ��ʱ��: " << duration.count() << " ����\n";

        writeCSV("output1.csv", data);

        data.clear();        
        pServ->Close();

    }
    catch (_com_error Err)
    {
        _tprintf(_T("Error:%s : 0x%x \n"), (TCHAR*)Err.Description(), Err.Error());
    }
    CoUninitialize();
    return 0;
}


int main() {
    //method1(); //6.4�����λ����ʱԼ150��
    //method2(); //�����̣߳�����ȡ1���λ����ʱԼ60��
    //method3(); //ͬ��
    //method4();  //���̶߳�ȡ1���λ����ʱ37��
    //method5();  //���Բɼ����㣬 2ms
    method6();  //�ʵ�ͻ�ƣ��ɼ�6���λ��1��
}