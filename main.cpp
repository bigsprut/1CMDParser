#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <shellapi.h>
#include <strsafe.h>
#include <string>
#include <vector>
#include <memory> 
#include "MDParser.h"

#define IDC_TABCONTROL    1000
#define IDC_TREEVIEW_OLE  1001 // ID для дерева файлов
#define IDC_EDITVIEW      1002
#define IDC_PATH_EDIT     1003
#define IDC_BROWSE_BTN    1004
#define IDC_LABEL_PATH    1005
#define IDC_TREEVIEW_META 1006 // ID для дерева метаданных

HINSTANCE g_hInst = NULL;
HWND g_hMainWnd = NULL;
HWND g_hTab = NULL;
HWND g_hTreeOLE = NULL;  // Дерево 1: Структура файла
HWND g_hTreeMeta = NULL; // Дерево 2: Метаданные
HWND g_hEdit = NULL;
HWND g_hPathEdit = NULL;
HWND g_hBrowseBtn = NULL;
HWND g_hPathLabel = NULL;

WCHAR g_szLastPath[MAX_PATH] = { 0 };
MDParser g_parser;

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
void ResizeControls(int width, int height);
void InitTabs(HWND hTab);
void SwitchTab(int index);
void OnBrowseFile();
void GetIniPath(WCHAR* buffer, size_t size);
void LoadSettings();
void SaveSettings();
void LoadAndParseFile(const WCHAR* path);

void FillTreeOLE(HTREEITEM hParent, const std::vector<OLEEntry>& entries);
void FillTreeMetadata(HTREEITEM hParent, std::shared_ptr<MdNode> node, int index);
void UpdateDetailView(LPNMTREEVIEWW pNM);

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int nCmdShow) {
    g_hInst = hInstance;

    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_TREEVIEW_CLASSES | ICC_TAB_CLASSES;
    InitCommonControlsEx(&icex);

    WNDCLASSEXW wcex = { 0 };
    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.hInstance = hInstance;
    wcex.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wcex.lpszClassName = L"OneCParserClass";

    if (!RegisterClassExW(&wcex)) return 1;

    g_hMainWnd = CreateWindowW(L"OneCParserClass", L"Парсер 1С 7.7 (Professional)", 
        WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN, CW_USEDEFAULT, 0, 900, 600, 
        NULL, NULL, hInstance, NULL);

    if (!g_hMainWnd) return 1;

    ShowWindow(g_hMainWnd, nCmdShow);
    UpdateWindow(g_hMainWnd);
    
    LoadSettings();
    if (lstrlenW(g_szLastPath) > 0) {
        SetWindowTextW(g_hPathEdit, g_szLastPath);
        LoadAndParseFile(g_szLastPath);
    }

    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    return (int)msg.wParam;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_CREATE:
        {
            HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
            HFONT hFixedFont = CreateFontW(16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, 
                RUSSIAN_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, 
                DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN, L"Courier New");

            g_hTab = CreateWindowExW(0, WC_TABCONTROLW, L"", 
                WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS, 
                0, 0, 0, 0, hWnd, (HMENU)IDC_TABCONTROL, g_hInst, NULL);
            SendMessage(g_hTab, WM_SETFONT, (WPARAM)hFont, 0);
            InitTabs(g_hTab);

            // 1. Создаем дерево OLE (видимое по умолчанию)
            g_hTreeOLE = CreateWindowExW(WS_EX_CLIENTEDGE, WC_TREEVIEWW, L"",
                WS_CHILD | WS_VISIBLE | WS_BORDER | TVS_HASLINES | TVS_LINESATROOT | TVS_HASBUTTONS | TVS_SHOWSELALWAYS,
                0, 0, 0, 0, hWnd, (HMENU)IDC_TREEVIEW_OLE, g_hInst, NULL);

            // 2. Создаем дерево Метаданных (скрытое по умолчанию)
            g_hTreeMeta = CreateWindowExW(WS_EX_CLIENTEDGE, WC_TREEVIEWW, L"",
                WS_CHILD | WS_BORDER | TVS_HASLINES | TVS_LINESATROOT | TVS_HASBUTTONS | TVS_SHOWSELALWAYS,
                0, 0, 0, 0, hWnd, (HMENU)IDC_TREEVIEW_META, g_hInst, NULL);
            
            g_hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | WS_HSCROLL | 
                ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL | ES_AUTOHSCROLL,
                0, 0, 0, 0, hWnd, (HMENU)IDC_EDITVIEW, g_hInst, NULL);
            
            SendMessage(g_hEdit, EM_SETLIMITTEXT, 0, 0);
            if (hFixedFont) SendMessage(g_hEdit, WM_SETFONT, (WPARAM)hFixedFont, 0);

            g_hPathLabel = CreateWindowW(L"STATIC", L"Путь к файлу 1cv7.md:",
                WS_CHILD | SS_SIMPLE, 0, 0, 0, 0, hWnd, (HMENU)IDC_LABEL_PATH, g_hInst, NULL);
            SendMessage(g_hPathLabel, WM_SETFONT, (WPARAM)hFont, 0);

            g_hPathEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                WS_CHILD | ES_AUTOHSCROLL | ES_READONLY, 
                0, 0, 0, 0, hWnd, (HMENU)IDC_PATH_EDIT, g_hInst, NULL);
            SendMessage(g_hPathEdit, WM_SETFONT, (WPARAM)hFont, 0);

            g_hBrowseBtn = CreateWindowW(L"BUTTON", L"Обзор...",
                WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, hWnd, (HMENU)IDC_BROWSE_BTN, g_hInst, NULL);
            SendMessage(g_hBrowseBtn, WM_SETFONT, (WPARAM)hFont, 0);

            SwitchTab(0);
        }
        break;

    case WM_SIZE:
        {
            int width = LOWORD(lParam);
            int height = HIWORD(lParam);
            SetWindowPos(g_hTab, NULL, 0, 0, width, height, SWP_NOZORDER);
            RECT rc;
            GetClientRect(g_hTab, &rc);
            TabCtrl_AdjustRect(g_hTab, FALSE, &rc);
            
            int treeW = (rc.right - rc.left) * 0.35; 
            
            // Изменяем размер обоих деревьев (даже скрытого), чтобы при переключении размер был верным
            SetWindowPos(g_hTreeOLE, HWND_TOP, rc.left, rc.top, treeW, rc.bottom - rc.top, 0);
            SetWindowPos(g_hTreeMeta, HWND_TOP, rc.left, rc.top, treeW, rc.bottom - rc.top, 0);
            
            SetWindowPos(g_hEdit, HWND_TOP, rc.left + treeW, rc.top, (rc.right - rc.left) - treeW, rc.bottom - rc.top, 0);

            SetWindowPos(g_hPathLabel, HWND_TOP, rc.left + 20, rc.top + 20, 300, 20, 0);
            SetWindowPos(g_hPathEdit, HWND_TOP, rc.left + 20, rc.top + 45, (rc.right - rc.left) - 140, 25, 0);
            SetWindowPos(g_hBrowseBtn, HWND_TOP, rc.right - 110, rc.top + 45, 90, 25, 0);
        }
        break;

    case WM_NOTIFY:
        {
            LPNMHDR pHdr = (LPNMHDR)lParam;
            if (pHdr->code == TCN_SELCHANGE) {
                int iPage = TabCtrl_GetCurSel(g_hTab);
                SwitchTab(iPage);
            }
            // Обрабатываем выбор в ЛЮБОМ из деревьев
            if ((pHdr->idFrom == IDC_TREEVIEW_OLE || pHdr->idFrom == IDC_TREEVIEW_META) 
                && pHdr->code == TVN_SELCHANGEDW) {
                
                LPNMTREEVIEWW pNMTV = (LPNMTREEVIEWW)lParam;
                UpdateDetailView(pNMTV);
            }
        }
        break;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDC_BROWSE_BTN) {
            OnBrowseFile();
        }
        break;

    case WM_DESTROY:
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProcW(hWnd, message, wParam, lParam);
    }
    return 0;
}

void InitTabs(HWND hTab) {
    TCITEMW tie;
    tie.mask = TCIF_TEXT;
    
    tie.pszText = (LPWSTR)L"Структура файла";
    TabCtrl_InsertItem(hTab, 0, &tie);
    
    tie.pszText = (LPWSTR)L"Метаданные";
    TabCtrl_InsertItem(hTab, 1, &tie);
    
    tie.pszText = (LPWSTR)L"Настройки";
    TabCtrl_InsertItem(hTab, 2, &tie);
}

void SwitchTab(int index) {
    // Управление видимостью контролов без удаления данных
    bool showOLE = (index == 0);
    bool showMeta = (index == 1);
    bool showSettings = (index == 2);
    bool showEdit = (index == 0 || index == 1);

    ShowWindow(g_hTreeOLE, showOLE ? SW_SHOW : SW_HIDE);
    ShowWindow(g_hTreeMeta, showMeta ? SW_SHOW : SW_HIDE);
    ShowWindow(g_hEdit, showEdit ? SW_SHOW : SW_HIDE);
    
    ShowWindow(g_hPathLabel, showSettings ? SW_SHOW : SW_HIDE);
    ShowWindow(g_hPathEdit, showSettings ? SW_SHOW : SW_HIDE);
    ShowWindow(g_hBrowseBtn, showSettings ? SW_SHOW : SW_HIDE);

    // При переключении восстанавливаем текст в Edit для активного элемента активного дерева
    SetWindowTextW(g_hEdit, L"");
    
    HTREEITEM hSel = NULL;
    HWND hActiveTree = NULL;
    int id = 0;

    if (showOLE) {
        hActiveTree = g_hTreeOLE;
        id = IDC_TREEVIEW_OLE;
    } else if (showMeta) {
        hActiveTree = g_hTreeMeta;
        id = IDC_TREEVIEW_META;
    }

    if (hActiveTree) {
        hSel = TreeView_GetSelection(hActiveTree);
        if (hSel) {
            // Эмулируем событие выбора, чтобы обновить Edit
            NM_TREEVIEWW nmtv = {0};
            nmtv.hdr.hwndFrom = hActiveTree;
            nmtv.hdr.idFrom = id;
            nmtv.hdr.code = TVN_SELCHANGEDW;
            nmtv.itemNew.hItem = hSel;
            nmtv.itemNew.mask = TVIF_PARAM;
            TreeView_GetItem(hActiveTree, &nmtv.itemNew); // Получаем lParam
            UpdateDetailView(&nmtv);
        }
    }
}

void OnBrowseFile() {
    OPENFILENAMEW ofn;       
    WCHAR szFile[MAX_PATH] = { 0 };
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g_hMainWnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    ofn.lpstrFilter = L"1C Configuration (*.md)\0*.md\0All Files\0*.*\0";
    ofn.nFilterIndex = 1;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

    if (GetOpenFileNameW(&ofn) == TRUE) {
        StringCchCopyW(g_szLastPath, MAX_PATH, szFile);
        SetWindowTextW(g_hPathEdit, g_szLastPath);
        SaveSettings();
        LoadAndParseFile(g_szLastPath);
        TabCtrl_SetCurSel(g_hTab, 0);
        SwitchTab(0);
    }
}

void GetIniPath(WCHAR* buffer, size_t size) {
    if (GetModuleFileNameW(NULL, buffer, (DWORD)size) == 0) {
        buffer[0] = 0; return;
    }
    WCHAR* pLastSlash = wcsrchr(buffer, L'\\');
    if (pLastSlash) *(pLastSlash + 1) = L'\0';
    StringCchCatW(buffer, size, L"parser.ini");
}

void LoadSettings() {
    WCHAR iniPath[MAX_PATH];
    GetIniPath(iniPath, MAX_PATH);
    GetPrivateProfileStringW(L"Settings", L"LastPath", L"", g_szLastPath, MAX_PATH, iniPath);
}

void SaveSettings() {
    WCHAR iniPath[MAX_PATH];
    GetIniPath(iniPath, MAX_PATH);
    WritePrivateProfileStringW(L"Settings", L"LastPath", g_szLastPath, iniPath);
}

void FillTreeOLE(HTREEITEM hParent, const std::vector<OLEEntry>& entries) {
    TVINSERTSTRUCTW tvis;
    tvis.hParent = hParent;
    tvis.hInsertAfter = TVI_LAST;
    tvis.item.mask = TVIF_TEXT | TVIF_CHILDREN | TVIF_PARAM; 

    for (const auto& entry : entries) {
        tvis.item.pszText = (LPWSTR)entry.name.c_str();
        tvis.item.cChildren = entry.isFolder ? 1 : 0;
        tvis.item.lParam = (LPARAM)&entry; 
        
        HTREEITEM hItem = TreeView_InsertItem(g_hTreeOLE, &tvis); // ВАЖНО: g_hTreeOLE
        
        if (entry.isFolder && !entry.children.empty()) {
            FillTreeOLE(hItem, entry.children);
        }
    }
}

void FillTreeMetadata(HTREEITEM hParent, std::shared_ptr<MdNode> node, int index) {
    if (!node) return;

    TVINSERTSTRUCTW tvis;
    tvis.hParent = hParent;
    tvis.hInsertAfter = TVI_LAST;
    tvis.item.mask = TVIF_TEXT | TVIF_CHILDREN | TVIF_PARAM; 

    std::wstring text;
    WCHAR buf[32];
    
    if (index >= 0) {
        StringCchPrintfW(buf, 32, L"[%d] ", index);
        text += buf;
    }

    if (!node->value.empty()) {
        int wlen = MultiByteToWideChar(1251, 0, node->value.c_str(), -1, NULL, 0);
        if (wlen > 0) {
            std::vector<WCHAR> wVal(wlen);
            MultiByteToWideChar(1251, 0, node->value.c_str(), -1, wVal.data(), wlen);
            text += L"\"";
            text += wVal.data();
            text += L"\"";
        }
    } else {
        bool hasNamedChild = false;
        if (!node->children.empty()) {
            auto firstChild = node->children[0];
            if (!firstChild->value.empty()) {
                int wlen = MultiByteToWideChar(1251, 0, firstChild->value.c_str(), -1, NULL, 0);
                if (wlen > 0) {
                    std::vector<WCHAR> wVal(wlen);
                    MultiByteToWideChar(1251, 0, firstChild->value.c_str(), -1, wVal.data(), wlen);
                    text += wVal.data(); 
                    hasNamedChild = true;
                }
            }
        }
        
        if (!hasNamedChild) {
            text += L"{...}";
        }
    }

    tvis.item.pszText = (LPWSTR)text.c_str();
    tvis.item.cChildren = node->children.empty() ? 0 : 1;
    tvis.item.lParam = (LPARAM)node.get(); 

    HTREEITEM hItem = TreeView_InsertItem(g_hTreeMeta, &tvis); // ВАЖНО: g_hTreeMeta

    int childIdx = 0;
    for (auto& child : node->children) {
        FillTreeMetadata(hItem, child, childIdx++);
    }
}

void LoadAndParseFile(const WCHAR* path) {
    // Очищаем оба дерева
    TreeView_DeleteAllItems(g_hTreeOLE);
    TreeView_DeleteAllItems(g_hTreeMeta);
    SetWindowTextW(g_hEdit, L"");
    
    if (g_parser.Open(path)) {
        // 1. Заполняем дерево OLE
        const auto& roots = g_parser.GetRootEntries();
        FillTreeOLE(TVI_ROOT, roots);

        // 2. Парсим метаданные и заполняем дерево Метаданных
        g_parser.ReadStreamText(L"Metadata\\Main MetaData Stream");
        auto parsedRoot = g_parser.GetParsedRoot();
        if (parsedRoot) {
            FillTreeMetadata(TVI_ROOT, parsedRoot, -1);
            HTREEITEM hRoot = TreeView_GetRoot(g_hTreeMeta);
            if (hRoot) TreeView_Expand(g_hTreeMeta, hRoot, TVE_EXPAND);
        }
    }
}

void UpdateDetailView(LPNMTREEVIEWW pNM) {
    // Определяем источник по ID контрола в заголовке сообщения
    if (pNM->hdr.idFrom == IDC_TREEVIEW_OLE) {
        OLEEntry* pEntry = (OLEEntry*)pNM->itemNew.lParam;
        if (pEntry) {
             std::wstring text;
             WCHAR buffer[256];
             text += L"=== СВОЙСТВА ОБЪЕКТА ===\r\n";
             StringCchPrintfW(buffer, 256, L"Имя:     %s\r\n", pEntry->name.c_str()); text += buffer;
             StringCchPrintfW(buffer, 256, L"Размер:  %u байт\r\n", pEntry->size); text += buffer;
             
             if (!pEntry->isFolder && pEntry->size > 0) {
                  text += L"\r\n=== СОДЕРЖИМОЕ ===\r\n";
                  text += g_parser.ReadStreamText(pEntry->fullPath);
             }
             SetWindowTextW(g_hEdit, text.c_str());
        }
    } 
    else if (pNM->hdr.idFrom == IDC_TREEVIEW_META) {
        MdNode* pNode = (MdNode*)pNM->itemNew.lParam;
        if (pNode) {
            std::wstring text = L"=== ФРАГМЕНТ МЕТАДАННЫХ ===\r\n";
            text += g_parser.DumpNodeToText(pNode);
            SetWindowTextW(g_hEdit, text.c_str());
        }
    }
}