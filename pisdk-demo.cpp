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

// 读取CSV文件
vector<string> readCSV(const string& filename)
{
    vector<string> data;
    ifstream file(filename);
    if (!file)
    {
        cout << "无法打开文件：" << filename << endl;
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

// 写入CSV文件
void writeCSV(const string& filename, const vector<vector<string>>& data)
{
    ofstream file(filename, std::ios::app);
    if (!file)
    {
        cout << "无法创建文件：" << filename << endl;
        return;
    }
    cout << "创建文件：" << filename << endl;

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

        std::cout << "代码执行时间: " << duration.count() << " 毫秒\n";
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

        std::cout << "代码执行时间: " << duration.count() << " 毫秒\n";
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


        // 计算sql查询的耗时
        auto start0 = std::chrono::high_resolution_clock::now();

        _PointListPtr pPointList = pServ->GetPoints((_bstr_t)"tag = 'XTEQ*'", 0); // 获取点位名为XTEQ开头的所有点位，大概64000个

        auto end0 = std::chrono::high_resolution_clock::now();
        auto duration0 = std::chrono::duration_cast<std::chrono::milliseconds>(end0 - start0);

        std::cout << "SQL查询时间: " << duration0.count() << " 毫秒\n";
        std::cout << "查询到数据共：" << pPointList->GetCount() << " 个\n";



        // 计算获取点位值的耗时
        auto start = std::chrono::high_resolution_clock::now();

        _variant_t var;
        VARIANT index;
        VariantInit(&index);
        index.vt = VT_INT;
        vector<vector<string>> datas;
        for (int i = 1; i <= pPointList->GetCount(); ++i) { //循环获取每个点位的值
            index.intVal = i;
            PIPointPtr pPt = pPointList->GetItem(&index);
            var = pPt->Data->GetSnapshot()->Value;
        }
        VariantClear(&index);
        auto end = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);

        std::cout << "代码执行时间: " << duration.count() << " 毫秒\n";
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
        // 从test1.csv中读取点位名称，存到一个列表中，1万个点位
        tags1 = readCSV("test1.csv");
        
        std::vector<string> tags2;
        tags2.reserve(10000);
        // 从test2.csv中读取点位名称，存到一个列表中，1万个点位
        tags2 = readCSV("test2.csv");


       PIPointsPtr pPoints = pServ->PIPoints;
        // 创建两个线程，共用一个PIPointsPtr去读取点位数据，每个线程分别读取10000个点位
        std::thread t1(GetTagsValueByPIPointsTh, pPoints, tags1);
        std::thread t2(GetTagsValueByPIPointsTh, pPoints, tags2);

        // 等待线程结束
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
        // 从test1.csv中读取点位名称，存到一个列表中，1万个点位
        tags1 = readCSV("test1.csv");

        std::vector<string> tags2;
        tags2.reserve(10000);
        // 从test2.csv中读取点位名称，存到一个列表中，1万个点位
        tags2 = readCSV("test2.csv");

        // 创建两个线程，每个线程有各自的PISDK实例去读取点位，每个线程分别读取10000个点位
        std::thread t1(GetTagsValueTh, tags1);
        std::thread t2(GetTagsValueTh, tags2);

        // 等待线程结束
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
        // 从test1.csv中读取点位名称，存到一个列表中，1万个点位
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

        std::cout << "代码执行时间: " << duration.count() << " 毫秒\n";
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
        // 从test.csv中读取点位名称，存到一个列表中，1万个点位
        tags = readCSV("test.csv");

        cout << "tags.size: " << tags.size() << endl;
        if (tags.empty())
            return -1;

        // 计算sql查询的耗时
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

        std::cout << "SQL查询时间: " << duration0.count() << " 毫秒\n";
        std::cout << "查询到数据共：" << pFinalPointList->GetCount() << " 个\n";


        // 计算获取点位值的耗时
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

        std::cout << "代码执行时间: " << duration.count() << " 毫秒\n";

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
    //method1(); //6.4万个点位，耗时约150秒
    //method2(); //两个线程，各读取1万点位，耗时约60秒
    //method3(); //同上
    //method4();  //单线程读取1万点位，耗时37秒
    //method5();  //测试采集单点， 2ms
    method6();  //质的突破，采集6万点位，1秒
}