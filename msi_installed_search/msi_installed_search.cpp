#include <windows.h>
#include <msi.h>
#include <msiquery.h>
#include <iostream>
#include <filesystem>
#include <comdef.h>
#include <thread> // For std::this_thread::sleep_for
#include <chrono> // For std::chrono::milliseconds
#include <sstream> // For std::wstringstream

void PrintMsiProperty(const wchar_t* databasePath, const char* propertyName) {
    MSIHANDLE hDatabase;

    UINT res = MsiOpenDatabase(databasePath, MSIDBOPEN_READONLY, &hDatabase);
    if (res != ERROR_SUCCESS) {
        std::cout << "Failed to open database: " << res << std::endl;
        return;
    }

    std::wstring query = L"SELECT `Value` FROM `Property` WHERE `Property`='" + std::wstring(propertyName, propertyName + strlen(propertyName)) + L"'";
    MSIHANDLE hView;
    res = MsiDatabaseOpenView(hDatabase, query.c_str(), &hView);
    if (res != ERROR_SUCCESS) {
        std::cout << "Failed to open view: " << res << std::endl;
        MsiCloseHandle(hDatabase);
        return;
    }

    res = MsiViewExecute(hView, 0);
    if (res != ERROR_SUCCESS) {
        std::cout << "Failed to execute view: " << res << std::endl;
        MsiCloseHandle(hView);
        MsiCloseHandle(hDatabase);
        return;
    }

    MSIHANDLE hRecord;
    res = MsiViewFetch(hView, &hRecord);
    if (res == ERROR_SUCCESS) {
        wchar_t value[256];
        DWORD valueSize = sizeof(value) / sizeof(value[0]);
        MsiRecordGetString(hRecord, 1, value, &valueSize);
        std::wcout << propertyName << L": " << value << std::endl;
        MsiCloseHandle(hRecord);
    }

    MsiCloseHandle(hView);
    MsiCloseHandle(hDatabase);
}

std::wstring FormatFilePath(const std::wstring& path) {
    std::wstringstream formattedPath;
    for (auto ch : path) {
        if (ch != L'\\') {
            formattedPath << ch;
        }
        else {
            formattedPath << L'\\'; // Keep single backslashes
        }
    }
    return formattedPath.str();
}

int main() {
    const std::wstring folderPath = L"C:\\Windows\\Installer";

    for (const auto& entry : std::filesystem::directory_iterator(folderPath)) {
        if (entry.path().extension() == L".msi") {
            std::wcout << L"-----------------------------" << std::endl;
            std::wcout << L"File: " << FormatFilePath(entry.path().wstring()) << std::endl;

            const wchar_t* filePath = entry.path().c_str();
            PrintMsiProperty(filePath, "Manufacturer");
            PrintMsiProperty(filePath, "ProductName");
            PrintMsiProperty(filePath, "ProductVersion");

            // Introduce a delay of 0.5 seconds (500 milliseconds)
            std::this_thread::sleep_for(std::chrono::milliseconds(500));
        }
    }

    return 0;
}