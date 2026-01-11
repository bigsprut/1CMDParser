# Makefile для NMAKE (Visual Studio)
# Целевая платформа: Windows Server 2003 (5.02)
# Кодировка: UTF-8

TARGET = parser.exe
SRC = main.cpp MDParser.cpp miniz.c
HEADERS = MDParser.h miniz.h

# Флаги компилятора
# /utf-8 - Важно для русского языка
# /MT - Статическая линковка (без DLL)
# /D "_CRT_SECURE_NO_WARNINGS" - чтобы miniz не ругался на fopen и strcpy
CPPFLAGS = /nologo /W3 /O2 /MT /EHsc /utf-8 \
           /D "WIN32" /D "_WINDOWS" /D "_UNICODE" /D "UNICODE" \
           /D "_WIN32_WINNT=0x0502" /D "_CRT_SECURE_NO_WARNINGS"

# Линковка с ole32.lib (для работы с файлами 1С)
LDFLAGS = /nologo /SUBSYSTEM:WINDOWS,5.02 \
          user32.lib kernel32.lib gdi32.lib comctl32.lib \
          comdlg32.lib ole32.lib shell32.lib advapi32.lib

all: $(TARGET)

$(TARGET): $(SRC) $(HEADERS)
	cl $(CPPFLAGS) $(SRC) /link $(LDFLAGS) /OUT:$(TARGET)

clean:
	del *.obj *.exe