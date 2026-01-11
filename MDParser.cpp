/*
 * Project: 1C 7.7 Configuration Parser
 * Author:  PrS <bigsprut@gmail.com>
 * GitHub:  https://github.com/bigsprut
 * License: MIT
 */
 
#include "MDParser.h"
#include "miniz.h" 
#include <comdef.h>
#include <sstream>
#include <vector>
#include <iomanip>
#include <algorithm>

// ============================================================================
// ХЕЛПЕРЫ: ZLIB / DECRYPT
// ============================================================================

bool IsZlib(const std::vector<char>& data) {
    if (data.size() < 2) return false;
    unsigned char b0 = (unsigned char)data[0];
    unsigned char b1 = (unsigned char)data[1];
    if (b0 != 0x78) return false;
    return (b1 == 0x9C || b1 == 0xDA || b1 == 0x01);
}

bool TryDecompress(std::vector<char>& data) {
    if (data.empty()) return false;
    size_t outSize = 0;
    // Используем функцию из miniz.c
    void* pDecomp = tinfl_decompress_mem_to_heap(data.data(), data.size(), &outSize, 0);
    if (pDecomp) {
        data.assign((char*)pDecomp, (char*)pDecomp + outSize);
        free(pDecomp);
        return true;
    }
    return false;
}

void ApplyDecrypt(std::vector<char>& data, const std::string& pass) {
    if (data.size() < 8) return;
    DWORD key = 0; 
    for (char c : pass) key = key * 4 + (unsigned char)c;
    
    DWORD rndSeed = 0; 
    memcpy(&rndSeed, &data[2], 4); 
    key ^= rndSeed;
    
    std::vector<char> output; 
    output.reserve(data.size() - 8);
    const DWORD LCG_MUL = 0x08088405; 
    const DWORD LCG_INC = 1;
    
    for (size_t i = 8; i < data.size(); ++i) {
        char c = (char)((unsigned char)data[i] ^ (unsigned char)(key & 0xFF));
        output.push_back(c);
        key = key * LCG_MUL + LCG_INC;
    }
    data = output;
}

int FindTextBrace(const std::vector<char>& data) {
    size_t limit = (std::min)((size_t)4096, data.size());
    for (size_t i = 0; i < limit; ++i) { 
        if (data[i] == '{') return (int)i; 
    }
    return -1;
}

// ============================================================================
// REALIZATION: MDParser
// ============================================================================

MDParser::MDParser() {
    CoInitialize(NULL);
}

MDParser::~MDParser() {
    Close();
    CoUninitialize();
}

void MDParser::Close() {
    rootEntries.clear();
    currentFilePath.clear();
    // Очистка структур парсера
    root.reset();
    objectIndex.clear();
    m_idToType.clear();
    m_fieldToRef.clear();
}

std::wstring MDParser::GetLastError() const {
    return lastError;
}

const std::vector<OLEEntry>& MDParser::GetRootEntries() const {
    return rootEntries;
}

std::shared_ptr<MdNode> MDParser::GetParsedRoot() const {
    return root;
}

bool MDParser::Open(const std::wstring& filePath) {
    Close();
    currentFilePath = filePath;

    IStorage* pRootStorage = NULL;
    
    HRESULT hr = StgOpenStorage(filePath.c_str(), NULL, 
        STGM_READ | STGM_SHARE_DENY_NONE | STGM_DIRECT, 
        NULL, 0, &pRootStorage);

    if (FAILED(hr)) {
        hr = StgOpenStorage(filePath.c_str(), NULL, 
            STGM_READ | STGM_SHARE_DENY_NONE | STGM_TRANSACTED, 
            NULL, 0, &pRootStorage);
    }

    if (FAILED(hr)) {
        lastError = L"Не удалось открыть файл. Код ошибки: " + std::to_wstring((long)hr);
        return false;
    }

    ReadStorage(pRootStorage, rootEntries, L"");
    pRootStorage->Release();
    return true;
}

void MDParser::ReadStorage(IStorage* pStorage, std::vector<OLEEntry>& targetList, const std::wstring& parentPath) {
    IEnumSTATSTG* pEnum = NULL;
    if (FAILED(pStorage->EnumElements(0, NULL, 0, &pEnum))) return;

    STATSTG stat;
    while (pEnum->Next(1, &stat, NULL) == S_OK) {
        OLEEntry entry;
        entry.name = stat.pwcsName;
        entry.size = stat.cbSize.LowPart;
        
        if (parentPath.empty()) entry.fullPath = entry.name;
        else entry.fullPath = parentPath + L"\\" + entry.name;

        CoTaskMemFree(stat.pwcsName);

        if (stat.type == STGTY_STORAGE) {
            entry.isFolder = true;
            IStorage* pSubStorage = NULL;
            if (pStorage->OpenStorage(entry.name.c_str(), NULL, 
                STGM_READ | STGM_SHARE_EXCLUSIVE, NULL, 0, &pSubStorage) == S_OK) {
                ReadStorage(pSubStorage, entry.children, entry.fullPath);
                pSubStorage->Release();
            }
        } else {
            entry.isFolder = false;
        }

        targetList.push_back(entry);
    }
    pEnum->Release();
}

// ============================================================================
// ПАРСИНГ СТРУКТУРЫ
// ============================================================================

void MDParser::SkipWhitespace(const char*& ptr) { 
    while (*ptr && (unsigned char)*ptr <= 32) ptr++; 
}

std::shared_ptr<MdNode> MDParser::ParseString(const char*& ptr) {
    SkipWhitespace(ptr); 

    auto node = std::make_shared<MdNode>();
    
    if (*ptr == '{') {
        ptr++; 
        SkipWhitespace(ptr); 
        
        while (*ptr && *ptr != '}') {
            auto child = ParseString(ptr);
            if (child) node->children.push_back(child);

            SkipWhitespace(ptr); 
            
            if (*ptr == ',') {
                ptr++; 
            } else if (*ptr != '}') {
                // Защита от зацикливания
            }
        }
        if (*ptr == '}') ptr++; 
    } 
    else if (*ptr == '"') {
        ptr++; 
        std::string val;
        while (*ptr) {
            if (*ptr == '"') { 
                if (*(ptr+1) == '"') { 
                    val += '"'; 
                    ptr += 2; 
                } else { 
                    ptr++; 
                    break; 
                } 
            } else { 
                val += *ptr++; 
            }
        }
        node->value = val;
    } else {
        // Чтение чисел
        std::string val;
        while (*ptr && *ptr != ',' && *ptr != '}' && (unsigned char)*ptr > 32) {
            val += *ptr++;
        }
        node->value = val;
    }
    
    return node;
}

void MDParser::ScanContainer(std::shared_ptr<MdNode> objectNode, const std::string& typePrefix) {
    if (objectNode->children.empty()) return;
    
    std::string objId = objectNode->children[0]->value;
    if (objId.empty()) return;
    
    m_idToType[objId] = typePrefix;
    objectIndex[objId] = objectNode; 
    
    for (auto& child : objectNode->children) {
        if (child->children.empty()) continue;
        
        std::string secName = child->children[0]->value;
        
        if (secName == "Head Fields" || secName == "Table Fields") {
            for (size_t i = 1; i < child->children.size(); ++i) {
                auto fieldNode = child->children[i];
                if (fieldNode->children.size() > 7) {
                    std::string fId = fieldNode->children[0]->value;
                    std::string fRef = fieldNode->children[7]->value;
                    
                    if (!fId.empty() && !fRef.empty() && fRef != "0") {
                        m_fieldToRef[fId] = fRef;
                    }
                }
            }
        }
    }
}

void MDParser::AnalyzeStructure() {
    if (!root) return;
    
    m_idToType.clear();
    m_fieldToRef.clear();
    objectIndex.clear();

    for (auto& section : root->children) {
        if (section->children.empty()) continue;
        
        std::string secType = section->children[0]->value;

        if (secType == "Documents") {
            for (size_t i = 1; i < section->children.size(); ++i) 
                ScanContainer(section->children[i], "DT");
        }
        else if (secType == "SbCnts") {
             for (size_t i = 1; i < section->children.size(); ++i) 
                ScanContainer(section->children[i], "SC");
        }
        else if (secType == "Registers") {
             for (size_t i = 1; i < section->children.size(); ++i) 
                ScanContainer(section->children[i], "RG");
        }
        else if (secType == "GenJrnlFldDef") {
             for (size_t i = 1; i < section->children.size(); ++i) {
                 auto fieldNode = section->children[i];
                 if (fieldNode->children.size() > 7) {
                    std::string fId = fieldNode->children[0]->value;
                    std::string fRef = fieldNode->children[7]->value;
                    if (!fId.empty() && !fRef.empty() && fRef != "0") {
                        m_fieldToRef[fId] = fRef;
                    }
                 }
             }
        }
    }
}

// Публичный метод для дампа узла
std::wstring MDParser::DumpNodeToText(const MdNode* node) {
    if (!node) return L"";
    std::wstringstream ss;
    DumpTreeToString(node, 0, ss);
    return ss.str();
}

// Рекурсивный вывод дерева (внутренний)
void MDParser::DumpTreeToString(const MdNode* node, int level, std::wstringstream& ss) {
    if (!node) return;

    for(int i=0; i<level; ++i) ss << L"  ";

    if (!node->value.empty()) {
        std::wstring wVal;
        int len = MultiByteToWideChar(1251, 0, node->value.c_str(), -1, NULL, 0);
        if (len > 0) {
            wVal.resize(len - 1);
            MultiByteToWideChar(1251, 0, node->value.c_str(), -1, &wVal[0], len - 1);
        }
        ss << L"\"" << wVal << L"\"";
    } else {
        ss << L"{...}";
    }

    if (!node->value.empty()) {
        auto itType = m_idToType.find(node->value);
        if (itType != m_idToType.end()) {
            std::wstring wType(itType->second.begin(), itType->second.end());
            ss << L" // Объект: " << wType;
        }
        auto itRef = m_fieldToRef.find(node->value);
        if (itRef != m_fieldToRef.end()) {
            std::wstring wRef(itRef->second.begin(), itRef->second.end());
            ss << L" // Ссылка на тип: " << wRef;
            auto itRefType = m_idToType.find(itRef->second);
            if (itRefType != m_idToType.end()) {
                std::wstring wRefType(itRefType->second.begin(), itRefType->second.end());
                ss << L" (" << wRefType << L")";
            }
        }
    }

    ss << L"\r\n";

    for (auto& child : node->children) {
        DumpTreeToString(child.get(), level + 1, ss);
    }
}

// ============================================================================
// ЧТЕНИЕ ПОТОКА И ИНТЕГРАЦИЯ
// ============================================================================

std::wstring MDParser::ReadStreamText(const std::wstring& fullPath) {
    if (fullPath.empty() || currentFilePath.empty()) return L"";

    std::vector<IStorage*> storageStack;
    IStorage* pRoot = NULL;
    
    HRESULT hr = StgOpenStorage(currentFilePath.c_str(), NULL, 
        STGM_READ | STGM_SHARE_DENY_NONE | STGM_DIRECT, NULL, 0, &pRoot);
    
    if (FAILED(hr)) {
        hr = StgOpenStorage(currentFilePath.c_str(), NULL, 
            STGM_READ | STGM_SHARE_DENY_NONE | STGM_TRANSACTED, NULL, 0, &pRoot);
    }

    if (FAILED(hr)) return L"Ошибка: Не удалось открыть файл-контейнер";
    storageStack.push_back(pRoot);

    std::wstring path = fullPath;
    IStorage* currStorage = pRoot;
    
    size_t pos = 0;
    while ((pos = path.find(L'\\')) != std::wstring::npos) {
        std::wstring folderName = path.substr(0, pos);
        path.erase(0, pos + 1);

        IStorage* nextStorage = NULL;
        hr = currStorage->OpenStorage(folderName.c_str(), NULL, 
            STGM_READ | STGM_SHARE_EXCLUSIVE, NULL, 0, &nextStorage);
        
        if (FAILED(hr)) {
            for (auto s : storageStack) s->Release();
            return L"Ошибка: Не удалось открыть папку [" + folderName + L"]";
        }
        currStorage = nextStorage;
        storageStack.push_back(currStorage);
    }

    IStream* pStream = NULL;
    hr = currStorage->OpenStream(path.c_str(), NULL, 
        STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &pStream);
    
    if (FAILED(hr)) {
        for (auto s : storageStack) s->Release();
        return L"Ошибка открытия потока";
    }

    std::wstring resultText = L"";
    STATSTG stat;
    hr = pStream->Stat(&stat, STATFLAG_NONAME);
    
    if (SUCCEEDED(hr)) {
        ULONG size = stat.cbSize.LowPart;
        if (size > 0) {
            std::vector<char> rawData(size);
            ULONG bytesRead = 0;
            hr = pStream->Read(rawData.data(), size, &bytesRead);
            
            if (SUCCEEDED(hr) && bytesRead > 0) {
                if (bytesRead < size) rawData.resize(bytesRead);

                bool readyToParse = false;
                
                // 1. ZLib?
                if (IsZlib(rawData)) {
                    if (TryDecompress(rawData)) readyToParse = true;
                }
                // 2. ZLib + Offset 8?
                if (!readyToParse && rawData.size() > 8) {
                    std::vector<char> copyOffset(rawData.begin() + 8, rawData.end());
                    if (IsZlib(copyOffset)) {
                        if (TryDecompress(copyOffset)) {
                            rawData = copyOffset;
                            readyToParse = true;
                        }
                    }
                }
                // 3. Encrypted?
                if (!readyToParse && rawData.size() > 8 && 
                    (unsigned char)rawData[0] == 0x25 && (unsigned char)rawData[1] == 0x77) {
                    std::vector<char> copyEnc = rawData;
                    ApplyDecrypt(copyEnc, ""); 
                    if (IsZlib(copyEnc)) {
                        if (TryDecompress(copyEnc)) {
                            rawData = copyEnc;
                            readyToParse = true;
                        }
                    }
                }
                // 4. Pure Text?
                if (!readyToParse && FindTextBrace(rawData) != -1) {
                    readyToParse = true;
                }

                if (readyToParse) {
                    // Если это поток метаданных, строим дерево
                    if (fullPath.find(L"Main MetaData Stream") != std::wstring::npos) {
                        int bracePos = FindTextBrace(rawData);
                        if (bracePos != -1) {
                            const char* ptr = rawData.data() + bracePos;
                            try {
                                root = ParseString(ptr);
                                AnalyzeStructure(); 
                                
                                std::wstringstream ss;
                                ss << L"=== СТРУКТУРА МЕТАДАННЫХ (PARSED) ===\r\n";
                                DumpTreeToString(root.get(), 0, ss);
                                resultText = ss.str();
                            } catch (...) {
                                resultText = L"Ошибка парсинга структуры";
                            }
                        } else {
                             resultText = L"Ошибка: данные распакованы, но не найден корневой элемент '{'";
                        }
                    } else {
                        // Для остальных потоков
                        int wlen = MultiByteToWideChar(1251, 0, rawData.data(), (int)rawData.size(), NULL, 0);
                        if (wlen > 0) {
                            resultText.resize(wlen);
                            MultiByteToWideChar(1251, 0, rawData.data(), (int)rawData.size(), &resultText[0], wlen);
                        }
                    }
                } else {
                    std::wstringstream ss;
                    ss << L"Неизвестный формат данных (RAW).\r\nHEX: ";
                    size_t dumpLen = (std::min)((size_t)32, rawData.size());
                    for (size_t i = 0; i < dumpLen; ++i) 
                         ss << std::hex << std::setw(2) << std::setfill(L'0') << (unsigned char)rawData[i] << L" ";
                    resultText = ss.str();
                }

            } else { resultText = L"Ошибка чтения (Read)"; }
        } else { resultText = L"<Пустой поток>"; }
    } else { resultText = L"Ошибка получения размера (Stat)"; }

    pStream->Release();
    for (int i = (int)storageStack.size() - 1; i >= 0; --i) storageStack[i]->Release();

    return resultText;
}