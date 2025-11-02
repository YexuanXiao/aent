#include <chrono>
#include <conio.h>
#include <exception>
#include <filesystem>
#include <fstream>
#include <memory>
#include <ranges>
#include <shellscalingapi.h>
#include <string>
#include <utility>
#include <windows.h>
#include <winevt.h>
#include <winrt/Windows.Data.Xml.Dom.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.UI.Notifications.h>
#include <winrt/base.h>
#include <winrt/windows.foundation.collections.h>

#pragma comment(lib, "runtimeobject.lib")
#pragma comment(lib, "wevtapi.lib")
#pragma comment(lib, "Shcore.lib")

#pragma comment(linker, "/subsystem:windows /entry:wmainCRTStartup")

namespace winrt
{
using namespace Windows::UI::Notifications;
using namespace Windows::Foundation;
using namespace Windows::Data::Xml::Dom;
} // namespace winrt

namespace bizwen
{

using namespace std::literals;

struct EventLog
{
    // System
    winrt::hstring systemTime;
    winrt::hstring UserId;

    // EventData
    winrt::hstring appName;
    winrt::hstring appVersion;
    winrt::hstring appTimeStamp;
    winrt::hstring moduleName;
    winrt::hstring moduleVersion;
    winrt::hstring moduleTimeStamp;
    winrt::hstring exceptionCode;
    winrt::hstring faultingOffset;
    winrt::hstring processId;
    winrt::hstring processCreationTime;
    winrt::hstring appPath;
    winrt::hstring modulePath;
    winrt::hstring integratorReportId;
    winrt::hstring packageFullName;
    winrt::hstring packageRelativeAppId;
};

void RegisterAumidForToast()
{
    HKEY hKey{};
    LONG result =
        ::RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\Classes\\AppUserModelId\\Application_Error_Notification_Tool",
                          0, nullptr, REG_OPTION_VOLATILE, KEY_WRITE, nullptr, &hKey, nullptr);

    if (result != ERROR_SUCCESS)
    {
        std::terminate();
    }

    const wchar_t displayName[] = L"Application Error Notification";
    result = ::RegSetValueExW(hKey, L"DisplayName", 0, REG_SZ, reinterpret_cast<const BYTE *>(displayName),
                              sizeof(displayName));

    if (result != ERROR_SUCCESS)
    {
        std::terminate();
    }

    if (RegCloseKey(hKey) != ERROR_SUCCESS)
    {
        std::terminate();
    }
}

void SendToNotificationCenter(std::wstring_view content)
{
    // <toast duration="short"><visual><binding template="ToastGeneric"><text>)
    // content
    // </text></binding></visual></toast>

    winrt::XmlDocument toastXml;

    auto toastElement = toastXml.CreateElement(L"toast"sv);
    toastElement.SetAttribute(L"duration"sv, L"short"sv);
    toastXml.AppendChild(toastElement);

    auto visualElement = toastXml.CreateElement(L"visual"sv);
    toastElement.AppendChild(visualElement);

    auto bindingElement = toastXml.CreateElement(L"binding"sv);
    bindingElement.SetAttribute(L"template"sv, L"ToastGeneric"sv);
    visualElement.AppendChild(bindingElement);

    auto textElement = toastXml.CreateElement(L"text"sv);
    textElement.InnerText(winrt::hstring(content));
    bindingElement.AppendChild(textElement);

    winrt::ToastNotification toast(toastXml);
    toast.ExpiresOnReboot(true);

    winrt::ToastNotificationManager::CreateToastNotifier(L"Application_Error_Notification_Tool"sv).Show(toast);
}

void CleanupRegistry()
{
    winrt::ToastNotificationManager::History().Clear(L"Application_Error_Notification_Tool"sv);

    auto errorCode =
        ::RegDeleteKeyW(HKEY_CURRENT_USER, L"Software\\Classes\\AppUserModelId\\Application_Error_Notification_Tool");
    if (errorCode != ERROR_SUCCESS && errorCode != ERROR_FILE_NOT_FOUND)
    {
        std::terminate();
    }
}

std::wstring FormatEventLog(const EventLog &eventLog, bool minimal)
{
    std::wstring output;
    output.reserve(800);

    auto appendIfNotEmpty = [&output, minimal](std::wstring_view fieldName, const winrt::hstring &fieldValue,
                                               bool fieldMinimal = false) {
        if (!fieldValue.empty() && !(minimal && !fieldMinimal))
        {
            output += fieldName;
            output += L": "sv;
            output += fieldValue;
            output += L'\n';
        }
    };
    appendIfNotEmpty(L"SystemTime"sv, eventLog.systemTime);
    appendIfNotEmpty(L"UserID"sv, eventLog.UserId);
    appendIfNotEmpty(L"AppName"sv, eventLog.appName, true);
    appendIfNotEmpty(L"AppVersion"sv, eventLog.appVersion);
    appendIfNotEmpty(L"AppTimeStamp"sv, eventLog.appTimeStamp);
    appendIfNotEmpty(L"ModuleName"sv, eventLog.moduleName, true);
    appendIfNotEmpty(L"ModuleVersion"sv, eventLog.moduleVersion);
    appendIfNotEmpty(L"ModuleTimeStamp"sv, eventLog.moduleTimeStamp);
    appendIfNotEmpty(L"ExceptionCode"sv, eventLog.exceptionCode, true);
    appendIfNotEmpty(L"FaultingOffset"sv, eventLog.faultingOffset);
    appendIfNotEmpty(L"ProcessId"sv, eventLog.processId);
    appendIfNotEmpty(L"ProcessCreationTime"sv, eventLog.processCreationTime);
    appendIfNotEmpty(L"AppPath"sv, eventLog.appPath);
    appendIfNotEmpty(L"ModulePath"sv, eventLog.modulePath);
    appendIfNotEmpty(L"IntegratorReportId"sv, eventLog.integratorReportId, true);
    appendIfNotEmpty(L"PackageFullName"sv, eventLog.packageFullName);
    appendIfNotEmpty(L"PackageRelativeAppId"sv, eventLog.packageRelativeAppId);

    return output;
}

void ParseEventLog(winrt::hstring xmlString, EventLog &eventLog) noexcept
{
    auto toXmlElement = [](auto &&node) { return node.template as<typename winrt::XmlElement>(); };

    winrt::XmlDocument doc;
    doc.LoadXml(xmlString);
    auto root = doc.DocumentElement();

    for (auto &&child : root.ChildNodes())
    {
        auto nodeName = child.NodeName();

        if (nodeName == L"System"sv)
        {
            for (auto &&element : child.ChildNodes() | std::views::transform(toXmlElement))
            {
                auto name = element.NodeName();

                if (name == L"TimeCreated"sv)
                {
                    eventLog.systemTime = element.GetAttribute(L"SystemTime"sv);
                }
                else if (name == L"Security"sv)
                {
                    eventLog.UserId = element.GetAttribute(L"UserID"sv);
                }
            }
        }
        else if (nodeName == L"EventData"sv)
        {
            for (auto &&element : child.ChildNodes() | std::views::transform(toXmlElement))
            {
                if (element.NodeName() == L"Data"sv)
                {
                    auto name = element.GetAttribute(L"Name"sv);
                    auto value = element.InnerText();

                    if (name == L"AppName"sv)
                        eventLog.appName = value;
                    else if (name == L"AppVersion"sv)
                        eventLog.appVersion = value;
                    else if (name == L"AppTimeStamp"sv)
                        eventLog.appTimeStamp = value;
                    else if (name == L"ModuleName"sv)
                        eventLog.moduleName = value;
                    else if (name == L"ModuleVersion"sv)
                        eventLog.moduleVersion = value;
                    else if (name == L"ModuleTimeStamp"sv)
                        eventLog.moduleTimeStamp = value;
                    else if (name == L"ExceptionCode"sv)
                        eventLog.exceptionCode = value;
                    else if (name == L"FaultingOffset"sv)
                        eventLog.faultingOffset = value;
                    else if (name == L"ProcessId"sv)
                        eventLog.processId = value;
                    else if (name == L"ProcessCreationTime"sv)
                        eventLog.processCreationTime = value;
                    else if (name == L"AppPath"sv)
                        eventLog.appPath = value;
                    else if (name == L"ModulePath"sv)
                        eventLog.modulePath = value;
                    else if (name == L"IntegratorReportId"sv)
                        eventLog.integratorReportId = value;
                    else if (name == L"PackageFullName"sv)
                        eventLog.packageFullName = value;
                    else if (name == L"PackageRelativeAppId"sv)
                        eventLog.packageRelativeAppId = value;
                }
            }
        }
    }
}

std::wstring WriteTempFile(std::wstring_view content)
{
    using namespace std::literals;
    auto path = std::filesystem::temp_directory_path();
    path /= std::to_string(std::chrono::steady_clock::now().time_since_epoch().count());
    path += L".txt"sv;

    std::wfstream fileStream(path.native(), std::ios::out | std::ios::binary);
    if (fileStream.is_open())
    {
        fileStream.write(content.data(), static_cast<std::streamsize>(content.size()));
    }
    else
    {
        std::terminate();
    }
    return path.native();
}

DWORD WINAPI MessageBoxThread(LPVOID lpParam)
{
    auto pMessage = static_cast<std::wstring *>(lpParam);
    if (::SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2) == nullptr)
    {
        std::terminate();
    }
    if (::MessageBoxW(nullptr, pMessage->c_str(), L"Application Error", MB_TOPMOST | MB_OK | MB_ICONERROR) == 0)
    {
        std::terminate();
    }
    delete pMessage;
    return 0;
}

void ShowMessageBoxAsync(std::wstring message)
{
    auto pMessage = new std::wstring(std::move(message));
    HANDLE hThread = ::CreateThread(nullptr, 0, MessageBoxThread, pMessage, 0, nullptr);
    if (!hThread)
    {
        std::terminate();
    }

    if (::CloseHandle(hThread) == 0)
    {
        std::terminate();
    }
}

void OpenNotepadWithContent(std::wstring content)
{
    auto tempFile = WriteTempFile(content);

    std::wstring parameter;
    parameter.reserve(content.size() + 2);
    parameter += L"\""sv;
    parameter += tempFile;
    parameter += L"\""sv;

    HINSTANCE hInst = ::ShellExecuteW(nullptr, L"open", L"notepad.exe", parameter.c_str(), nullptr, SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(hInst) <= 32)
    {
        std::terminate();
    }
}

void OpenPowerShellWithContent(std::wstring content)
{
    auto tempFile = WriteTempFile(content);
    auto prefix = L"-NoExit Get-Content -Path \""sv;
    auto postfix = L"\""sv;
    std::wstring parameter;
    parameter.reserve(content.size() + prefix.size() + postfix.size());
    parameter += prefix;
    parameter += tempFile;
    parameter += postfix;

    HINSTANCE hInst = ::ShellExecuteW(nullptr, L"open", L"powershell.exe", parameter.c_str(), nullptr, SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(hInst) <= 32)
    {
        std::terminate();
    }
}

enum class PrintMethod
{
    console = 0b1,
    messagebox = 0b10,
    notepad = 0b100,
    powershell = 0b1000,
    notification = 0b10000
};

constexpr PrintMethod &operator|=(PrintMethod &a, PrintMethod b) noexcept
{
    return a = static_cast<PrintMethod>(std::to_underlying(a) | std::to_underlying(b));
}

constexpr bool operator&(PrintMethod a, PrintMethod b) noexcept
{
    return (std::to_underlying(a) & std::to_underlying(b)) != 0;
}

enum class PrintStyle
{
    text,
    xml
};

enum class RunMode
{
    console,
    silent,
    kill,
    help
};

void ParseArguments(std::wstring_view arg, PrintMethod &method, PrintStyle &style, RunMode &mode)
{
    if (arg == L"-messagebox"sv)
    {
        method |= PrintMethod::messagebox;
    }
    else if (arg == L"-notepad"sv)
    {
        method |= PrintMethod::notepad;
    }
    else if (arg == L"-powershell"sv)
    {
        method |= PrintMethod::powershell;
    }
    else if (arg == L"-notification"sv)
    {
        method |= PrintMethod::notification;
    }
    else if (arg == L"-text"sv)
    {
        style = PrintStyle::text;
    }
    else if (arg == L"-xml"sv)
    {
        style = PrintStyle::xml;
    }
    else if (arg == L"-console"sv)
    {
        mode = RunMode::console;
    }
    else if (arg == L"-silent"sv)
    {
        mode = RunMode::silent;
    }
    else if (arg == L"-kill")
    {
        mode = RunMode::kill;
    }
    else if (arg == L"-help")
    {
        mode = RunMode::help;
    }
    else
    {
        std::terminate();
    }
}

void WriteContentConsole(std::wstring_view str)
{
    if (auto handle = ::GetStdHandle(STD_OUTPUT_HANDLE))
    {
        DWORD outSize;
        if (!::WriteConsoleW(handle, str.data(), static_cast<DWORD>(str.size()), &outSize, nullptr))
        {
            std::terminate();
        }
    }
}

EVT_HANDLE SubscribeEvent(HANDLE event)
{
    auto hSubscription = ::EvtSubscribe(nullptr, event, L"Application", L"*[System[(Level=2)]]", nullptr, nullptr,
                                        nullptr, EvtSubscribeToFutureEvents);
    if (hSubscription == nullptr)
    {
        std::terminate();
    }
    return hSubscription;
}

std::wstring PrintEvent(EVT_HANDLE hEvent)
{
    DWORD dwBufferSize = 0;
    DWORD dwBufferUsed = 0;
    DWORD dwPropertyCount = 0;

    ::EvtRender(nullptr, hEvent, EvtRenderEventXml, dwBufferSize, nullptr, &dwBufferUsed, &dwPropertyCount);

    if (::GetLastError() == ERROR_INSUFFICIENT_BUFFER)
    {
        dwBufferSize = dwBufferUsed;
        std::wstring renderedContent;
        renderedContent.resize_and_overwrite(dwBufferSize / sizeof(wchar_t), [&](wchar_t *data, size_t size) {
            if (::EvtRender(nullptr, hEvent, EvtRenderEventXml, dwBufferSize, data, &dwBufferUsed, &dwPropertyCount))
            {
                return dwBufferUsed / sizeof(wchar_t) - 1; // NB: null terminator
            }
            else
            {
                std::terminate();
            }
        });
        return renderedContent;
    }
    else
    {
        std::terminate();
    }
}

void DispatchOutput(PrintMethod method, const std::wstring &content)
{
    if (method & PrintMethod::console)
    {
        WriteContentConsole(content);
    }
    if (method & PrintMethod::messagebox)
    {
        ShowMessageBoxAsync(content);
    }
    if (method & PrintMethod::notepad)
    {
        OpenNotepadWithContent(content);
    }
    if (method & PrintMethod::powershell)
    {
        OpenPowerShellWithContent(content);
    }
    if (method & PrintMethod::notification)
    {
        SendToNotificationCenter(content);
    }
}

void DispatchOutput(PrintMethod method, const EventLog &eventLog)
{
    if (method & PrintMethod::console)
    {
        auto content = FormatEventLog(eventLog, false);
        WriteContentConsole(content);
    }
    if (method & PrintMethod::messagebox)
    {
        auto content = FormatEventLog(eventLog, false);
        ShowMessageBoxAsync(content);
    }
    if (method & PrintMethod::notepad)
    {
        auto content = FormatEventLog(eventLog, false);
        OpenNotepadWithContent(content);
    }
    if (method & PrintMethod::powershell)
    {
        auto content = FormatEventLog(eventLog, false);
        OpenPowerShellWithContent(content);
    }
    if (method & PrintMethod::notification)
    {
        auto content = FormatEventLog(eventLog, true);
        SendToNotificationCenter(content);
    }
}

void EnumerateResults(EVT_HANDLE hResults, PrintMethod method, PrintStyle style)
{
    while (true)
    {
        EVT_HANDLE hEvent;
        DWORD dwReturned = 0;
        if (EvtNext(hResults, 1, &hEvent, INFINITE, 0, &dwReturned))
        {
            auto xml = PrintEvent(hEvent);
            if (style == PrintStyle::text)
            {
                EventLog eventLog;
                ParseEventLog(winrt::hstring(xml), eventLog);
                DispatchOutput(method, eventLog);
            }
            else
            {
                DispatchOutput(method, xml);
            }
            EvtClose(hEvent);
        }
        else if (GetLastError() == ERROR_NO_MORE_ITEMS)
        {
            return;
        }
        else
        {
            std::terminate();
        }
    }
}

bool IsKeyEvent(HANDLE hStdIn)
{
    INPUT_RECORD record;
    DWORD dwRecordsRead = 0;

    if (::ReadConsoleInput(hStdIn, &record, 1, &dwRecordsRead))
    {
        return record.EventType == KEY_EVENT;
    }
    else
    {
        std::terminate();
    }
}

void WaitOnEvent(PrintMethod method, PrintStyle style)
{
    HANDLE aWaitHandles[2];

    aWaitHandles[0] = ::CreateEventW(nullptr, TRUE, FALSE, L"Application_Error_Notification_Tool");

    if (aWaitHandles[0] == nullptr)
    {
        std::terminate();
    }

    aWaitHandles[1] = ::CreateEventW(nullptr, TRUE, TRUE, nullptr);

    if (aWaitHandles[1] == nullptr)
    {
        std::terminate();
    }

    auto hSubscription = SubscribeEvent(aWaitHandles[1]);

    while (true)
    {
        DWORD dwWait = WaitForMultipleObjects(2, aWaitHandles, FALSE, INFINITE);

        if (dwWait == WAIT_OBJECT_0) // Kill event
        {
            break;
        }
        else if (dwWait == WAIT_OBJECT_0 + 1) // Query results
        {
            EnumerateResults(hSubscription, method, style);

            ResetEvent(aWaitHandles[1]);
        }
        else
        {
            std::terminate();
        }
    }

    EvtClose(hSubscription);
    CloseHandle(aWaitHandles[0]);
    CloseHandle(aWaitHandles[1]);
}

void WaitOnConsole(PrintMethod method, PrintStyle style)
{
    HANDLE aWaitHandles[2];

    aWaitHandles[0] = GetStdHandle(STD_INPUT_HANDLE);

    if (aWaitHandles[0] == INVALID_HANDLE_VALUE || aWaitHandles[0] == nullptr)
    {
        std::terminate();
    }

    aWaitHandles[1] = CreateEventW(nullptr, TRUE, TRUE, nullptr);

    if (aWaitHandles[1] == nullptr)
    {
        std::terminate();
    }

    auto hSubscription = SubscribeEvent(aWaitHandles[1]);

    while (true)
    {
        DWORD dwWait = WaitForMultipleObjects(2, aWaitHandles, FALSE, INFINITE);

        if (dwWait == WAIT_OBJECT_0) // Console input
        {
            if (IsKeyEvent(aWaitHandles[0]))
                break;
        }
        else if (dwWait == WAIT_OBJECT_0 + 1) // Query results
        {
            EnumerateResults(hSubscription, method, style);

            ResetEvent(aWaitHandles[1]);
            WriteContentConsole(L"Waiting, press any key to exit.\n"sv);
        }
        else
        {
            std::terminate();
        }
    }

    EvtClose(hSubscription);

    CloseHandle(aWaitHandles[1]);
}

bool TryAttachConsole()
{
    if (::AttachConsole(ATTACH_PARENT_PROCESS) == 0)
    {
        auto errorCode = ::GetLastError();
        if (errorCode == ERROR_ACCESS_DENIED)
        {
            // attach
            return true;
        }
        else if (errorCode == ERROR_INVALID_HANDLE)
        {
            // no console
            return false;
        }
        else
        {
            return false;
        }
    }
    return true;
}

} // namespace bizwen

int wmain(int argc, wchar_t **argv)
{

    using namespace std::literals;
    auto usage = LR"(
Application Error Notification Tool
Copyright (c) 2025 YexuanXiao
usage: apperrnotitool.exe [options...]
options:
    -console     : Start with current console
    -silent      : Start without attaching to the console
    -kill        : Kill the running service
    -help        : Print the help message
    -messagebox  : Show info via MessageBox
    -notepad     : Show info via Windows Notepad
    -powershell  : Show info via Windows PowerShell
    -notification: Show info via Notification Center
    -text        : Output info as text
    -xml         : Output info as unformatted XML
)"sv;

    bizwen::PrintMethod method{};
    bizwen::PrintStyle style{};
    bizwen::RunMode mode{};

    for (int i = 1; i != argc; ++i)
    {
        ParseArguments(argv[i], method, style, mode);
    }

    if (method & bizwen::PrintMethod::notification)
    {
        bizwen::RegisterAumidForToast();
    }

    if (mode == bizwen::RunMode::help)
    {
        bizwen::TryAttachConsole();
        bizwen::WriteContentConsole(usage);
    }
    else if (mode == bizwen::RunMode::kill)
    {
        if (auto hEvent = ::OpenEventW(EVENT_MODIFY_STATE, FALSE, L"Application_Error_Notification_Tool");
            hEvent != nullptr)
        {
            if (::SetEvent(hEvent) == FALSE)
            {
                std::terminate();
            }
            if (::CloseHandle(hEvent) == 0)
            {
                std::terminate();
            }
			bizwen::TryAttachConsole();
            bizwen::WriteContentConsole(L"The background service has been terminated.\n"sv);
        }
        else if (::GetLastError() == ERROR_FILE_NOT_FOUND)
        {
			bizwen::TryAttachConsole();
            bizwen::WriteContentConsole(L"The background service is not running.\n"sv);
        }
        else
        {
            std::terminate();
        }
    }
    else if (mode == bizwen::RunMode::silent)
    {
        if (auto hEvent = ::OpenEventW(EVENT_MODIFY_STATE, FALSE, L"Application_Error_Notification_Tool");
            hEvent != nullptr)
        {
            if (::CloseHandle(hEvent) == 0)
            {
                std::terminate();
            }
        }
        else
        {
            WaitOnEvent(method, style);
        }
    }
    else
    {
        method |= bizwen::PrintMethod::console;
        if (auto hEvent = ::OpenEventW(EVENT_MODIFY_STATE, FALSE, L"Application_Error_Notification_Tool");
            hEvent != nullptr)
        {
            if (::CloseHandle(hEvent) == 0)
            {
                std::terminate();
            }
            bizwen::TryAttachConsole();
            bizwen::WriteContentConsole(L"The service is already running, please kill it first.\n"sv);
        }
        else
        {
			 // NB
            if (bizwen::TryAttachConsole())
            {
                WaitOnConsole(method, style);
            }
            else
            {
                std::terminate();
            }
        }
    }

    if (method & bizwen::PrintMethod::notification)
    {
        bizwen::CleanupRegistry();
    }

    return 0;
}
