/*
 * Project: 1C 7.7 Configuration Parser
 * Author:  PrS <bigsprut@gmail.com>
 * GitHub:  https://github.com/bigsprut
 * License: MIT
 */
 
#pragma once
#include <windows.h>
#include <string>
#include <vector>
#include <map>
#include <memory>
#include <ole2.h>
#include <sstream>

// Структура узла метаданных (дерево)
struct MdNode {
    std::string value; // Значение узла (в кодировке 1251)
    std::vector<std::shared_ptr<MdNode>> children;
};

// Структура для отображения в TreeView (файловая система OLE)
struct OLEEntry {
    std::wstring name;
    std::wstring fullPath;
    bool isFolder;
    unsigned long size;
    std::vector<OLEEntry> children;
};

class MDParser {
public:
    MDParser();
    ~MDParser();

    bool Open(const std::wstring& filePath);
    void Close();
    
    // Получить корневые элементы OLE (файлы/папки)
    const std::vector<OLEEntry>& GetRootEntries() const;
    
    // Получить корень распарсенного дерева метаданных (для GUI)
    std::shared_ptr<MdNode> GetParsedRoot() const;

    std::wstring GetLastError() const;

    // Читает поток, определяет формат (ZLib/Crypt), парсит структуру и возвращает текст
    std::wstring ReadStreamText(const std::wstring& entryParams);

    // Генерирует текстовый дамп конкретного узла и его детей (для GUI)
    std::wstring DumpNodeToText(const MdNode* node);

private:
    std::wstring lastError;
    std::wstring currentFilePath;
    std::vector<OLEEntry> rootEntries;

    // === Структуры парсера метаданных ===
    std::shared_ptr<MdNode> root;
    std::map<std::string, std::shared_ptr<MdNode>> objectIndex; 
    std::map<std::string, std::string> m_idToType;   // ID объекта -> Тип
    std::map<std::string, std::string> m_fieldToRef; // ID поля -> ID типа назначения

    // === Внутренние методы ===
    void ReadStorage(IStorage* pStorage, std::vector<OLEEntry>& targetList, const std::wstring& parentPath);
    
    // Парсинг строки 1С {"...", ...}
    std::shared_ptr<MdNode> ParseString(const char*& ptr);
    void SkipWhitespace(const char*& ptr);
    
    // Анализ структуры после парсинга (заполнение карт типов)
    void AnalyzeStructure();
    void ScanContainer(std::shared_ptr<MdNode> objectNode, const std::string& typePrefix);
    
    // Рекурсивный вывод дерева в поток (принимает сырой указатель для удобства)
    void DumpTreeToString(const MdNode* node, int level, std::wstringstream& ss);
};