#pragma comment(linker, "/subsystem:windows")

#include <windows.h>
#include <errhandlingapi.h>

int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine,
                    _In_ int nShowCmd)
{
    RaiseException(EXCEPTION_ACCESS_VIOLATION, 0, 0, nullptr);
}
