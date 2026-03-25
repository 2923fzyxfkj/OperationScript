#include <windows.h>
#include <shellapi.h>
#include <gdiplus.h>

#include <algorithm>
#include <chrono>
#include <cmath>
#include <cctype>
#include <ctime>
#include <filesystem>
#include <fstream>
#include <functional>
#include <iomanip>
#include <iostream>
#include <map>
#include <sstream>
#include <stdexcept>
#include <string>
#include <vector>

namespace {

using namespace Gdiplus;

struct CommandInfo {
    std::string category;
    std::string name;
    std::string syntax;
    std::string description;
    bool implemented;
};

struct Point {
    int x = 0;
    int y = 0;
};

struct ScriptLine {
    std::string text;
    int lineNumber = 0;
};

struct FunctionInfo {
    std::string name;
    std::vector<std::string> params;
    size_t startIndex = 0;
    size_t endIndex = 0;
};

struct WindowSearchContext {
    std::wstring target;
    HWND hwnd = nullptr;
};

struct BreakSignal {};
struct ContinueSignal {};

struct ImageData {
    int width = 0;
    int height = 0;
    std::vector<unsigned char> pixels;
};

std::wstring utf8ToWide(const std::string& text) {
    if (text.empty()) {
        return L"";
    }
    int size = MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, nullptr, 0);
    std::wstring result(size - 1, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, result.data(), size);
    return result;
}

std::string wideToUtf8(const std::wstring& text) {
    if (text.empty()) {
        return "";
    }
    int size = WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string result(size - 1, '\0');
    WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1, result.data(), size, nullptr, nullptr);
    return result;
}

std::string trim(const std::string& value) {
    const auto begin = value.find_first_not_of(" \t\r\n");
    if (begin == std::string::npos) {
        return "";
    }
    const auto end = value.find_last_not_of(" \t\r\n");
    return value.substr(begin, end - begin + 1);
}

std::string stripBom(const std::string& value) {
    if (value.size() >= 3 &&
        static_cast<unsigned char>(value[0]) == 0xEF &&
        static_cast<unsigned char>(value[1]) == 0xBB &&
        static_cast<unsigned char>(value[2]) == 0xBF) {
        return value.substr(3);
    }
    return value;
}

bool startsWith(const std::string& text, const std::string& prefix) {
    return text.rfind(prefix, 0) == 0;
}

std::vector<std::string> tokenize(const std::string& line) {
    std::vector<std::string> tokens;
    std::string current;
    bool inQuote = false;
    for (size_t i = 0; i < line.size(); ++i) {
        const char ch = line[i];
        if (ch == '"') {
            inQuote = !inQuote;
            continue;
        }
        if (!inQuote && std::isspace(static_cast<unsigned char>(ch))) {
            if (!current.empty()) {
                tokens.push_back(current);
                current.clear();
            }
            continue;
        }
        current.push_back(ch);
    }
    if (!current.empty()) {
        tokens.push_back(current);
    }
    return tokens;
}

std::string join(const std::vector<std::string>& parts, size_t start) {
    std::ostringstream out;
    for (size_t i = start; i < parts.size(); ++i) {
        if (i > start) {
            out << ' ';
        }
        out << parts[i];
    }
    return out.str();
}

double parseNumber(const std::string& text) {
    try {
        return std::stod(text);
    } catch (...) {
        throw std::runtime_error("无法解析数字: " + text);
    }
}

Point parsePoint(const std::string& text) {
    if (text.size() < 5 || text.front() != '(' || text.back() != ')') {
        throw std::runtime_error("坐标格式应为 (x,y): " + text);
    }
    const std::string inner = text.substr(1, text.size() - 2);
    const auto comma = inner.find(',');
    if (comma == std::string::npos) {
        throw std::runtime_error("坐标格式应为 (x,y): " + text);
    }
    Point point;
    point.x = static_cast<int>(parseNumber(trim(inner.substr(0, comma))));
    point.y = static_cast<int>(parseNumber(trim(inner.substr(comma + 1))));
    return point;
}

double easeProgress(double t, const std::string& mode) {
    if (mode == "线性") {
        return t;
    }
    if (mode == "缓入") {
        return t * t * t;
    }
    if (mode == "缓出") {
        const double inv = 1.0 - t;
        return 1.0 - inv * inv * inv;
    }
    if (mode == "缓入缓出") {
        return t < 0.5 ? 4.0 * t * t * t : 1.0 - std::pow(-2.0 * t + 2.0, 3.0) / 2.0;
    }
    throw std::runtime_error("未知滑动类型: " + mode);
}

void sendMouseDown(bool left) {
    INPUT input{};
    input.type = INPUT_MOUSE;
    input.mi.dwFlags = left ? MOUSEEVENTF_LEFTDOWN : MOUSEEVENTF_RIGHTDOWN;
    SendInput(1, &input, sizeof(INPUT));
}

void sleepMs(int milliseconds) {
    Sleep(milliseconds < 0 ? 0 : milliseconds);
}

void sleepSeconds(double seconds) {
    sleepMs(static_cast<int>(seconds * 1000.0));
}

void sendMouseUp(bool left) {
    INPUT input{};
    input.type = INPUT_MOUSE;
    input.mi.dwFlags = left ? MOUSEEVENTF_LEFTUP : MOUSEEVENTF_RIGHTUP;
    SendInput(1, &input, sizeof(INPUT));
}

void sendMouseClick(bool left) {
    sendMouseDown(left);
    sleepMs(25);
    sendMouseUp(left);
}

void sendUnicodeChar(wchar_t ch) {
    INPUT inputs[2]{};
    inputs[0].type = INPUT_KEYBOARD;
    inputs[0].ki.dwFlags = KEYEVENTF_UNICODE;
    inputs[0].ki.wScan = ch;
    inputs[1].type = INPUT_KEYBOARD;
    inputs[1].ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
    inputs[1].ki.wScan = ch;
    SendInput(2, inputs, sizeof(INPUT));
}

void sendMouseWheel(int delta) {
    INPUT input{};
    input.type = INPUT_MOUSE;
    input.mi.dwFlags = MOUSEEVENTF_WHEEL;
    input.mi.mouseData = static_cast<DWORD>(delta);
    SendInput(1, &input, sizeof(INPUT));
}

void sendVirtualKey(WORD key) {
    INPUT inputs[2]{};
    inputs[0].type = INPUT_KEYBOARD;
    inputs[0].ki.wVk = key;
    inputs[1].type = INPUT_KEYBOARD;
    inputs[1].ki.wVk = key;
    inputs[1].ki.dwFlags = KEYEVENTF_KEYUP;
    SendInput(2, inputs, sizeof(INPUT));
}

BOOL CALLBACK enumWindowByTitle(HWND hwnd, LPARAM lParam) {
    auto* context = reinterpret_cast<WindowSearchContext*>(lParam);
    if (!IsWindowVisible(hwnd)) {
        return TRUE;
    }
    const int length = GetWindowTextLengthW(hwnd);
    if (length <= 0) {
        return TRUE;
    }
    std::wstring title(length, L'\0');
    GetWindowTextW(hwnd, title.data(), length + 1);
    if (title.find(context->target) != std::wstring::npos) {
        context->hwnd = hwnd;
        return FALSE;
    }
    return TRUE;
}

HWND findWindowByTitle(const std::wstring& title) {
    WindowSearchContext context{title, nullptr};
    EnumWindows(enumWindowByTitle, reinterpret_cast<LPARAM>(&context));
    return context.hwnd;
}

std::string normalizeColorText(std::string color) {
    color = trim(color);
    if (startsWith(color, "#")) {
        color.erase(color.begin());
    }
    if (startsWith(color, "0x") || startsWith(color, "0X")) {
        color = color.substr(2);
    }
    std::transform(color.begin(), color.end(), color.begin(), [](unsigned char ch) {
        return static_cast<char>(std::toupper(ch));
    });
    return color;
}

std::string colorToHex(COLORREF color) {
    char buffer[16]{};
    std::snprintf(buffer, sizeof(buffer), "%02X%02X%02X", GetRValue(color), GetGValue(color), GetBValue(color));
    return buffer;
}

ULONG_PTR ensureGdiplusStarted() {
    static ULONG_PTR token = 0;
    static bool started = false;
    if (!started) {
        GdiplusStartupInput input;
        const Status status = GdiplusStartup(&token, &input, nullptr);
        if (status != Ok) {
            throw std::runtime_error("GDI+ 初始化失败");
        }
        started = true;
    }
    return token;
}

ImageData captureScreenImage(int x, int y, int width, int height) {
    HDC screenDc = GetDC(nullptr);
    HDC memDc = CreateCompatibleDC(screenDc);
    HBITMAP bitmap = CreateCompatibleBitmap(screenDc, width, height);
    HGDIOBJ oldObject = SelectObject(memDc, bitmap);
    BitBlt(memDc, 0, 0, width, height, screenDc, x, y, SRCCOPY);

    BITMAPINFOHEADER bi{};
    bi.biSize = sizeof(BITMAPINFOHEADER);
    bi.biWidth = width;
    bi.biHeight = -height;
    bi.biPlanes = 1;
    bi.biBitCount = 32;
    bi.biCompression = BI_RGB;

    ImageData image;
    image.width = width;
    image.height = height;
    image.pixels.resize(static_cast<size_t>(width) * static_cast<size_t>(height) * 4);
    if (!GetDIBits(memDc, bitmap, 0, height, image.pixels.data(),
                   reinterpret_cast<BITMAPINFO*>(&bi), DIB_RGB_COLORS)) {
        SelectObject(memDc, oldObject);
        DeleteObject(bitmap);
        DeleteDC(memDc);
        ReleaseDC(nullptr, screenDc);
        throw std::runtime_error("读取屏幕像素失败");
    }

    SelectObject(memDc, oldObject);
    DeleteObject(bitmap);
    DeleteDC(memDc);
    ReleaseDC(nullptr, screenDc);
    return image;
}

ImageData loadImageFile(const std::filesystem::path& path) {
    ensureGdiplusStarted();
    std::wstring widePath = utf8ToWide(path.string());
    Bitmap source(widePath.c_str());
    if (source.GetLastStatus() != Ok) {
        throw std::runtime_error("无法加载图片: " + path.string());
    }

    Bitmap converted(source.GetWidth(), source.GetHeight(), PixelFormat32bppARGB);
    Graphics graphics(&converted);
    graphics.DrawImage(&source, 0, 0, source.GetWidth(), source.GetHeight());

    Rect rect(0, 0, converted.GetWidth(), converted.GetHeight());
    BitmapData data{};
    if (converted.LockBits(&rect, ImageLockModeRead, PixelFormat32bppARGB, &data) != Ok) {
        throw std::runtime_error("读取图片像素失败: " + path.string());
    }

    ImageData image;
    image.width = converted.GetWidth();
    image.height = converted.GetHeight();
    image.pixels.resize(static_cast<size_t>(image.width) * static_cast<size_t>(image.height) * 4);

    for (int y = 0; y < image.height; ++y) {
        const unsigned char* row = reinterpret_cast<const unsigned char*>(data.Scan0) + static_cast<size_t>(y) * data.Stride;
        std::copy(row, row + static_cast<size_t>(image.width) * 4,
                  image.pixels.begin() + static_cast<size_t>(y) * image.width * 4);
    }

    converted.UnlockBits(&data);
    return image;
}

std::string findImageOnScreen(const ImageData& needle) {
    const int screenX = GetSystemMetrics(SM_XVIRTUALSCREEN);
    const int screenY = GetSystemMetrics(SM_YVIRTUALSCREEN);
    const int screenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
    const int screenHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);

    if (needle.width <= 0 || needle.height <= 0 ||
        needle.width > screenWidth || needle.height > screenHeight) {
        return "false";
    }

    const ImageData haystack = captureScreenImage(screenX, screenY, screenWidth, screenHeight);
    const size_t needleRowBytes = static_cast<size_t>(needle.width) * 4;
    const size_t haystackRowBytes = static_cast<size_t>(haystack.width) * 4;

    for (int y = 0; y <= haystack.height - needle.height; ++y) {
        for (int x = 0; x <= haystack.width - needle.width; ++x) {
            bool match = true;
            for (int yy = 0; yy < needle.height && match; ++yy) {
                const size_t haystackOffset = static_cast<size_t>(y + yy) * haystackRowBytes + static_cast<size_t>(x) * 4;
                const size_t needleOffset = static_cast<size_t>(yy) * needleRowBytes;
                if (!std::equal(
                        needle.pixels.begin() + static_cast<std::ptrdiff_t>(needleOffset),
                        needle.pixels.begin() + static_cast<std::ptrdiff_t>(needleOffset + needleRowBytes),
                        haystack.pixels.begin() + static_cast<std::ptrdiff_t>(haystackOffset))) {
                    match = false;
                }
            }
            if (match) {
                return "(" + std::to_string(screenX + x) + "," + std::to_string(screenY + y) + ")";
            }
        }
    }
    return "false";
}

void saveBitmapToBmpFile(HBITMAP bitmap, HDC hdc, const std::filesystem::path& path) {
    BITMAP bmp{};
    GetObject(bitmap, sizeof(BITMAP), &bmp);

    BITMAPINFOHEADER bi{};
    bi.biSize = sizeof(BITMAPINFOHEADER);
    bi.biWidth = bmp.bmWidth;
    bi.biHeight = -bmp.bmHeight;
    bi.biPlanes = 1;
    bi.biBitCount = 32;
    bi.biCompression = BI_RGB;

    const DWORD bitmapSize = static_cast<DWORD>(bmp.bmWidth * bmp.bmHeight * 4);
    std::vector<unsigned char> pixels(bitmapSize);
    if (!GetDIBits(hdc, bitmap, 0, bmp.bmHeight, pixels.data(),
                   reinterpret_cast<BITMAPINFO*>(&bi), DIB_RGB_COLORS)) {
        throw std::runtime_error("读取截图像素失败");
    }

    BITMAPFILEHEADER bfh{};
    bfh.bfType = 0x4D42;
    bfh.bfOffBits = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER);
    bfh.bfSize = bfh.bfOffBits + bitmapSize;

    std::ofstream output(path, std::ios::binary);
    if (!output) {
        throw std::runtime_error("无法写入截图文件: " + path.string());
    }
    output.write(reinterpret_cast<const char*>(&bfh), sizeof(bfh));
    output.write(reinterpret_cast<const char*>(&bi), sizeof(bi));
    output.write(reinterpret_cast<const char*>(pixels.data()), bitmapSize);
}

std::pair<Point, Point> parsePointRange(const std::string& text) {
    const size_t arrow = text.find("->");
    if (arrow == std::string::npos) {
        throw std::runtime_error("区域格式应为 (x1,y1)->(x2,y2)");
    }
    return {parsePoint(text.substr(0, arrow)), parsePoint(text.substr(arrow + 2))};
}

std::string escapePowerShellSingleQuoted(const std::string& text) {
    std::string result;
    for (char ch : text) {
        result.push_back(ch);
        if (ch == '\'') {
            result.push_back('\'');
        }
    }
    return result;
}

std::string downloadUrlTextViaPowerShell(const std::string& url) {
    const std::filesystem::path tempPath =
        std::filesystem::temp_directory_path() / ("operation_script_web_" + std::to_string(GetTickCount()) + ".txt");
    const std::string command =
        "powershell -Command \""
        "$ProgressPreference='SilentlyContinue'; "
        "Invoke-WebRequest -UseBasicParsing -Uri '" + escapePowerShellSingleQuoted(url) +
        "' | Select-Object -ExpandProperty Content | Set-Content -Encoding UTF8 '" +
        escapePowerShellSingleQuoted(tempPath.string()) + "'\"";
    if (std::system(command.c_str()) != 0) {
        throw std::runtime_error("网页请求失败: " + url);
    }
    std::ifstream input(tempPath, std::ios::binary);
    if (!input) {
        throw std::runtime_error("读取网页结果失败");
    }
    std::ostringstream buffer;
    buffer << input.rdbuf();
    input.close();
    std::filesystem::remove(tempPath);
    std::string content = buffer.str();
    content = stripBom(content);
    return content;
}

std::string requestUrlViaPowerShell(const std::string& method, const std::string& url, const std::string& body) {
    const std::filesystem::path tempPath =
        std::filesystem::temp_directory_path() / ("operation_script_req_" + std::to_string(GetTickCount()) + ".txt");
    std::string command =
        "powershell -Command \""
        "$ProgressPreference='SilentlyContinue'; "
        "$r = Invoke-WebRequest -UseBasicParsing -Method " + method +
        " -Uri '" + escapePowerShellSingleQuoted(url) + "'";
    if (!body.empty()) {
        command += " -Body '" + escapePowerShellSingleQuoted(body) + "'";
    }
    command += " ; $r.Content | Set-Content -Encoding UTF8 '" +
               escapePowerShellSingleQuoted(tempPath.string()) + "'\"";
    if (std::system(command.c_str()) != 0) {
        throw std::runtime_error("网络请求失败: " + url);
    }

    std::ifstream input(tempPath, std::ios::binary);
    if (!input) {
        throw std::runtime_error("读取请求结果失败");
    }
    std::ostringstream buffer;
    buffer << input.rdbuf();
    input.close();
    std::filesystem::remove(tempPath);
    return stripBom(buffer.str());
}

bool detectNetworkViaPowerShell() {
    const std::string command =
        "powershell -Command \""
        "$ProgressPreference='SilentlyContinue'; "
        "$targets = @("
        "  'https://www.msftconnecttest.com/connecttest.txt', "
        "  'http://www.msftconnecttest.com/connecttest.txt', "
        "  'https://www.qq.com/', "
        "  'https://www.baidu.com/'"
        "); "
        "foreach ($uri in $targets) { "
        "  try { "
        "    Invoke-WebRequest -UseBasicParsing -Method Get -TimeoutSec 5 -Uri $uri | Out-Null; "
        "    exit 0; "
        "  } catch { } "
        "} "
        "try { "
        "  [System.Net.Dns]::GetHostEntry('www.baidu.com') | Out-Null; "
        "  exit 0; "
        "} catch { "
        "  exit 1; "
        "}\"";
    return std::system(command.c_str()) == 0;
}

std::string runOcrViaPowerShell(const std::filesystem::path& imagePath, const std::filesystem::path& scriptPath) {
    const std::filesystem::path outputPath =
        std::filesystem::temp_directory_path() / ("operation_script_ocr_" + std::to_string(GetTickCount()) + ".txt");
    const std::string command =
        "powershell -ExecutionPolicy Bypass -File \"" + scriptPath.string() +
        "\" -ImagePath \"" + imagePath.string() +
        "\" -OutputPath \"" + outputPath.string() + "\"";
    if (std::system(command.c_str()) != 0) {
        throw std::runtime_error("OCR 执行失败");
    }
    std::ifstream input(outputPath, std::ios::binary);
    if (!input) {
        throw std::runtime_error("OCR 输出读取失败");
    }
    std::ostringstream buffer;
    buffer << input.rdbuf();
    input.close();
    std::filesystem::remove(outputPath);
    return stripBom(buffer.str());
}

std::filesystem::path getExecutableDirectory() {
    char buffer[MAX_PATH]{};
    GetModuleFileNameA(nullptr, buffer, MAX_PATH);
    return std::filesystem::path(buffer).parent_path();
}

std::string formatNow(const char* format) {
    std::time_t now = std::time(nullptr);
    std::tm local = *std::localtime(&now);
    char buffer[128]{};
    std::strftime(buffer, sizeof(buffer), format, &local);
    return buffer;
}

std::string toDisplayNumber(double value) {
    std::ostringstream out;
    out << std::setprecision(15) << value;
    std::string result = out.str();
    if (result.find('.') != std::string::npos) {
        while (!result.empty() && result.back() == '0') {
            result.pop_back();
        }
        if (!result.empty() && result.back() == '.') {
            result.pop_back();
        }
    }
    return result.empty() ? "0" : result;
}

std::string extractBracketValue(const std::string& text, const std::string& prefix) {
    if (!startsWith(text, prefix + "[") || text.back() != ']') {
        throw std::runtime_error("参数格式错误，需要 " + prefix + "[值]");
    }
    return text.substr(prefix.size() + 1, text.size() - prefix.size() - 2);
}

std::string extractAnyBracketValue(const std::string& text, const std::vector<std::string>& prefixes) {
    for (const auto& prefix : prefixes) {
        if (startsWith(text, prefix + "[") && !text.empty() && text.back() == ']') {
            return text.substr(prefix.size() + 1, text.size() - prefix.size() - 2);
        }
    }
    throw std::runtime_error("参数格式错误");
}

std::pair<std::string, std::string> splitNameAndBracket(const std::string& text) {
    const size_t left = text.find('[');
    if (left == std::string::npos) {
        return {text, ""};
    }
    if (text.back() != ']') {
        throw std::runtime_error("参数格式错误，缺少 ]");
    }
    const std::string name = text.substr(0, left);
    if (name.empty()) {
        throw std::runtime_error("参数格式错误，变量名不能为空");
    }
    return {name, text.substr(left + 1, text.size() - left - 2)};
}

std::vector<std::string> parseOptionList(const std::string& text) {
    std::string source = text;
    if (!source.empty() && source.front() == '[' && source.back() == ']') {
        source = source.substr(1, source.size() - 2);
    }

    std::vector<std::string> result;
    std::string current;
    for (char ch : source) {
        if (ch == '"') {
            continue;
        }
        if (ch == ',') {
            current = trim(current);
            if (!current.empty()) {
                result.push_back(current);
            }
            current.clear();
            continue;
        }
        current.push_back(ch);
    }
    current = trim(current);
    if (!current.empty()) {
        result.push_back(current);
    }
    if (result.empty()) {
        throw std::runtime_error("选项列表不能为空");
    }
    return result;
}

std::vector<std::string> splitArguments(const std::string& text) {
    std::vector<std::string> result;
    std::string current;
    bool inQuote = false;
    int bracketDepth = 0;
    int parenDepth = 0;

    for (char ch : text) {
        if (ch == '"') {
            inQuote = !inQuote;
            current.push_back(ch);
            continue;
        }
        if (!inQuote) {
            if (ch == '[') ++bracketDepth;
            if (ch == ']') --bracketDepth;
            if (ch == '(') ++parenDepth;
            if (ch == ')') --parenDepth;
            if (ch == ',' && bracketDepth == 0 && parenDepth == 0) {
                result.push_back(trim(current));
                current.clear();
                continue;
            }
        }
        current.push_back(ch);
    }

    if (!trim(current).empty()) {
        result.push_back(trim(current));
    }
    if (result.size() == 1 && result[0].empty()) {
        result.clear();
    }
    return result;
}

std::vector<std::string> splitListItems(const std::string& text) {
    std::string source = trim(text);
    if (source.empty()) {
        return {};
    }
    if (source.front() == '[' && source.back() == ']') {
        source = source.substr(1, source.size() - 2);
    }
    std::vector<std::string> result;
    std::string current;
    bool inQuote = false;
    for (char ch : source) {
        if (ch == '"') {
            inQuote = !inQuote;
            continue;
        }
        if (!inQuote && (ch == ',' || ch == ';')) {
            current = trim(current);
            if (!current.empty()) {
                result.push_back(current);
            }
            current.clear();
            continue;
        }
        current.push_back(ch);
    }
    current = trim(current);
    if (!current.empty()) {
        result.push_back(current);
    }
    return result;
}

bool parseFunctionSignature(const std::string& text, std::string& name, std::vector<std::string>& args) {
    const std::string trimmed = trim(text);
    const size_t left = trimmed.find('(');
    const size_t right = trimmed.rfind(')');
    if (left == std::string::npos || right == std::string::npos || right <= left) {
        return false;
    }
    name = trim(trimmed.substr(0, left));
    if (name.empty()) {
        return false;
    }
    args = splitArguments(trimmed.substr(left + 1, right - left - 1));
    return true;
}

struct PopupDialogState {
    Point origin;
    std::wstring message;
    std::wstring buttonText;
    std::vector<std::wstring> options;
    bool showEdit = false;
    bool showSelect = false;
    bool confirmed = false;
    std::wstring inputValue;
    int selectedIndex = -1;
    HWND editHandle = nullptr;
};

LRESULT CALLBACK popupDialogProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    PopupDialogState* state = reinterpret_cast<PopupDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    switch (message) {
    case WM_NCCREATE: {
        CREATESTRUCTW* create = reinterpret_cast<CREATESTRUCTW*>(lParam);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
        return TRUE;
    }
    case WM_CREATE: {
        state = reinterpret_cast<PopupDialogState*>(reinterpret_cast<CREATESTRUCTW*>(lParam)->lpCreateParams);
        HFONT font = static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
        HWND label = CreateWindowW(
            L"STATIC", state->message.c_str(),
            WS_CHILD | WS_VISIBLE,
            16, 16, 320, 48,
            hwnd, nullptr, nullptr, nullptr);
        SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);

        int currentY = 72;
        if (state->showEdit) {
            state->editHandle = CreateWindowExW(
                WS_EX_CLIENTEDGE, L"EDIT", L"",
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
                16, currentY, 320, 24,
                hwnd, reinterpret_cast<HMENU>(1001), nullptr, nullptr);
            SendMessageW(state->editHandle, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
            currentY += 36;
        }
        if (state->showSelect) {
            const int buttonWidth = 90;
            const int buttonHeight = 28;
            const int gap = 8;
            const int totalWidth = static_cast<int>(state->options.size()) * buttonWidth +
                                   static_cast<int>(state->options.size() - 1) * gap;
            int startX = std::max(16, (360 - totalWidth) / 2);
            for (size_t i = 0; i < state->options.size(); ++i) {
                HWND button = CreateWindowW(
                    L"BUTTON", state->options[i].c_str(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | (i == 0 ? BS_DEFPUSHBUTTON : 0),
                    startX + static_cast<int>(i) * (buttonWidth + gap), currentY, buttonWidth, buttonHeight,
                    hwnd, reinterpret_cast<HMENU>(2000 + static_cast<int>(i)), nullptr, nullptr);
                SendMessageW(button, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
            }
        } else {
            HWND button = CreateWindowW(
                L"BUTTON", state->buttonText.c_str(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
                136, currentY, 80, 28,
                hwnd, reinterpret_cast<HMENU>(1003), nullptr, nullptr);
            SendMessageW(button, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
        }
        return 0;
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == 1003) {
            if (state != nullptr) {
                if (state->showEdit && state->editHandle != nullptr) {
                    const int length = GetWindowTextLengthW(state->editHandle);
                    std::wstring value(length, L'\0');
                    GetWindowTextW(state->editHandle, value.data(), length + 1);
                    state->inputValue = value;
                }
                state->confirmed = true;
            }
            DestroyWindow(hwnd);
            return 0;
        }
        if (LOWORD(wParam) >= 2000 && LOWORD(wParam) < 3000 && state != nullptr) {
            state->confirmed = true;
            state->selectedIndex = static_cast<int>(LOWORD(wParam) - 2000);
            DestroyWindow(hwnd);
            return 0;
        }
        break;
    case WM_CLOSE:
        DestroyWindow(hwnd);
        return 0;
    default:
        break;
    }
    return DefWindowProcW(hwnd, message, wParam, lParam);
}

bool runPopupDialog(PopupDialogState& state, const std::wstring& title) {
    static bool classRegistered = false;
    const wchar_t* className = L"OperationScriptPopupDialog";
    if (!classRegistered) {
        WNDCLASSW wc{};
        wc.lpfnWndProc = popupDialogProc;
        wc.hInstance = GetModuleHandleW(nullptr);
        wc.lpszClassName = className;
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
        RegisterClassW(&wc);
        classRegistered = true;
    }

    const int width = 360;
    const int height = state.showEdit || state.showSelect ? 170 : 130;
    HWND hwnd = CreateWindowExW(
        WS_EX_TOPMOST,
        className,
        title.c_str(),
        WS_CAPTION | WS_SYSMENU,
        state.origin.x, state.origin.y, width, height,
        nullptr, nullptr, GetModuleHandleW(nullptr), &state);

    if (hwnd == nullptr) {
        throw std::runtime_error("弹窗创建失败");
    }

    ShowWindow(hwnd, SW_SHOW);
    UpdateWindow(hwnd);

    MSG msg{};
    while (IsWindow(hwnd)) {
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
            if (msg.message == WM_QUIT) {
                return state.confirmed;
            }
            if (!IsDialogMessageW(hwnd, &msg)) {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
        }
        sleepMs(10);
    }
    return state.confirmed;
}

std::vector<CommandInfo> buildCommandCatalog() {
    return {
        {"鼠标", "点击", "点击 左键|右键 单点|双击|多击|长按 (x,y) [次数或秒数]", "在指定坐标执行鼠标点击。", true},
        {"鼠标", "移动鼠标", "移动鼠标 (x,y)", "把鼠标移动到指定坐标。", true},
        {"鼠标", "滑动", "滑动 左键|右键 秒数 线性|缓入|缓出|缓入缓出 (x1,y1)->(x2,y2)", "按住鼠标并沿轨迹移动。", true},
        {"鼠标", "滚轮", "滚轮 上|下 次数", "滚动鼠标滚轮。", true},
        {"鼠标", "拖到", "拖到 左键|右键 (x1,y1)->(x2,y2)", "执行拖拽。", true},
        {"鼠标", "鼠标回中", "鼠标回中", "把鼠标移动到屏幕中心。", true},
        {"鼠标", "鼠标隐藏", "鼠标隐藏", "隐藏鼠标指针。", false},
        {"鼠标", "鼠标显示", "鼠标显示", "显示鼠标指针。", false},
        {"鼠标", "锁定鼠标", "锁定鼠标 区域 (x1,y1)->(x2,y2)", "限制鼠标活动区域。", false},
        {"鼠标", "解锁鼠标", "解锁鼠标", "解除鼠标区域限制。", false},

        {"键盘", "按键", "按键 键名", "敲击一个按键。", true},
        {"键盘", "长按键", "长按键 键名 秒数", "长按某个键。", false},
        {"键盘", "松开键", "松开键 键名", "松开某个键。", false},
        {"键盘", "组合键", "组合键 Ctrl+Shift+S", "按下组合键。", false},
        {"键盘", "输入文本", "输入文本 \"内容\"", "模拟键盘输入文本。", true},
        {"键盘", "清空输入", "清空输入", "发送清空输入框动作。", false},
        {"键盘", "复制", "复制", "发送 Ctrl+C。", false},
        {"键盘", "粘贴", "粘贴", "发送 Ctrl+V。", false},
        {"键盘", "剪切", "剪切", "发送 Ctrl+X。", false},
        {"键盘", "撤销", "撤销", "发送 Ctrl+Z。", false},
        {"键盘", "重做", "重做", "发送 Ctrl+Y。", false},
        {"键盘", "回车", "回车", "发送 Enter。", true},
        {"键盘", "退格", "退格", "发送 Backspace。", true},
        {"键盘", "删除键", "删除键", "发送 Delete。", true},
        {"键盘", "截图键", "截图键", "发送 PrintScreen。", false},

        {"系统", "执行", "执行 0|1 命令文本", "0=cmd，1=powershell。", true},
        {"系统", "打开", "打开 网址|路径|程序", "等价于系统 open/start。", true},
        {"系统", "关闭程序", "关闭程序 程序名", "关闭指定进程。", false},
        {"系统", "重启程序", "重启程序 程序名", "重启某程序。", false},
        {"系统", "等待", "等待 秒数", "暂停脚本。", true},
        {"系统", "暂停", "暂停", "等待用户确认后继续。", true},
        {"系统", "退出脚本", "退出脚本 [退出码]", "结束当前脚本。", true},
        {"系统", "关机", "关机", "关闭系统。", false},
        {"系统", "重启电脑", "重启电脑", "重启系统。", false},
        {"系统", "锁屏", "锁屏", "锁定系统。", false},
        {"系统", "注销", "注销", "注销当前用户。", false},
        {"系统", "播放声音", "播放声音 文件路径", "播放音频。", false},
        {"系统", "蜂鸣", "蜂鸣 频率 毫秒", "播放蜂鸣音。", true},
        {"系统", "设置音量", "设置音量 0-100", "设置系统音量。", false},
        {"系统", "静音", "静音", "系统静音。", false},
        {"系统", "取消静音", "取消静音", "恢复声音。", false},
        {"系统", "亮度", "亮度 0-100", "设置屏幕亮度。", false},
        {"系统", "截图", "截图 保存到 路径", "截取屏幕。", false},
        {"系统", "录屏", "录屏 开始|停止 [路径]", "控制录屏。", false},
        {"系统", "通知", "通知 标题 内容", "弹出系统通知。", false},

        {"窗口", "激活窗口", "激活窗口 标题", "激活匹配标题的窗口。", true},
        {"窗口", "关闭窗口", "关闭窗口 标题", "关闭指定窗口。", true},
        {"窗口", "最小化窗口", "最小化窗口 标题", "最小化窗口。", true},
        {"窗口", "最大化窗口", "最大化窗口 标题", "最大化窗口。", true},
        {"窗口", "还原窗口", "还原窗口 标题", "还原窗口。", true},
        {"窗口", "移动窗口", "移动窗口 标题 (x,y)", "移动窗口位置。", false},
        {"窗口", "调整窗口", "调整窗口 标题 宽 高", "调整窗口大小。", false},
        {"窗口", "置顶窗口", "置顶窗口 标题", "窗口置顶。", false},
        {"窗口", "取消置顶", "取消置顶 标题", "取消置顶。", false},
        {"窗口", "读取窗口文本", "读取窗口文本 标题 到 变量名", "抓取窗口文本。", false},

        {"文件", "复制文件", "复制文件 源 到 目标", "复制文件。", true},
        {"文件", "移动文件", "移动文件 源 到 目标", "移动文件。", true},
        {"文件", "删除文件", "删除文件 路径", "删除文件。", true},
        {"文件", "创建文件", "创建文件 路径", "创建空文件。", true},
        {"文件", "写入文件", "写入文件 路径 内容", "覆盖写入文件。", true},
        {"文件", "追加文件", "追加文件 路径 内容", "追加写入文件。", true},
        {"文件", "读取文件", "读取文件 路径 到 变量名", "读取文本文件。", true},
        {"文件", "创建目录", "创建目录 路径", "创建文件夹。", true},
        {"文件", "删除目录", "删除目录 路径", "删除文件夹。", true},
        {"文件", "列出目录", "列出目录 路径 到 变量名", "读取目录项。", true},
        {"文件", "文件存在", "文件存在 路径 到 变量名", "检查文件是否存在。", true},
        {"文件", "路径存在", "路径存在 路径 到 变量名", "检查路径是否存在。", true},
        {"文件", "文件大小", "文件大小 路径 到 变量名", "获取文件大小。", true},
        {"文件", "改名文件", "改名文件 旧路径 新路径", "重命名文件。", true},
        {"文件", "解压文件", "解压文件 压缩包 到 目录", "解压压缩文件。", false},

        {"变量", "变量", "变量 名称 [= 值]", "定义或修改变量，缺省值为 0。", true},
        {"变量", "加", "加 名称 数值", "变量加法，推荐改用 名称 = 名称 加 数值。", true},
        {"变量", "减", "减 名称 数值", "变量减法，推荐改用 名称 = 名称 减 数值。", true},
        {"变量", "乘", "乘 名称 数值", "变量乘法，推荐改用 名称 = 名称 乘 数值。", true},
        {"变量", "除", "除 名称 数值", "变量除法，推荐改用 名称 = 名称 除 数值。", true},
        {"变量", "连接", "连接 名称 值", "字符串拼接。", true},
        {"变量", "删除变量", "删除变量 名称", "删除变量。", true},
        {"变量", "读取环境", "读取环境 变量名 到 名称", "读取环境变量。", true},
        {"变量", "设置环境", "设置环境 变量名 值", "设置环境变量。", true},
        {"变量", "转整数", "转整数 名称", "把变量转换为整数。", true},
        {"变量", "转小数", "转小数 名称", "把变量转换为浮点数。", true},
        {"变量", "转文本", "转文本 名称", "把变量转换为文本。", true},

        {"流程", "如果", "如果 左值 运算符 右值", "支持单行则语法，也支持缩进块。", true},
        {"流程", "否则如果", "否则如果 左值 运算符 右值", "与如果配套的分支判断。", true},
        {"流程", "否则", "否则", "与如果配套，推荐使用缩进块。", true},
        {"流程", "当", "当 左值 运算符 右值", "支持单行时语法，也支持缩进块。", true},
        {"流程", "重复", "重复 次数", "支持单行命令，也支持缩进块。", true},
        {"流程", "遍历", "遍历 变量 在 列表", "支持单行命令，也支持缩进块。", true},
        {"流程", "跳出", "跳出", "跳出循环。", true},
        {"流程", "继续", "继续", "继续下一次循环。", true},
        {"流程", "标签", "标签 名称", "定义跳转标签。", true},
        {"流程", "去", "去 标签名", "跳转到标签。", true},
        {"流程", "调用", "调用 脚本路径", "调用另一个 os 脚本。", true},
        {"流程", "返回", "返回 [值]", "从函数返回值。", true},
        {"流程", "函数", "函数 名称(参数1,参数2)", "定义函数，推荐使用缩进块。", true},
        {"流程", "结束", "结束", "旧式代码块结束标记，已兼容保留。", true},

        {"交互", "输出", "输出 内容", "输出文本到控制台。", true},
        {"交互", "换行", "换行", "输出空行。", true},
        {"交互", "清屏", "清屏", "清空控制台。", true},
        {"交互", "标题", "标题 内容", "设置控制台标题。", true},
        {"交互", "询问", "询问 提示 到 变量名", "读取用户输入。", true},
        {"交互", "确认", "确认 提示 到 变量名", "获取是/否。", true},
        {"交互", "选择", "选择 提示 选项 到 变量名", "获取用户选项。", false},
        {"交互", "确定弹窗", "确定弹窗 (x,y) 消息 变量名[OK文字]", "按确定写入 true，关闭则写入 false。", true},
        {"交互", "询问弹窗", "询问弹窗 (x,y) 消息 变量名[OK文字]", "按确定写入输入内容，关闭则写入 false。", true},
        {"交互", "选择弹窗", "选择弹窗 (x,y) 消息 变量名[选项1,选项2,...]", "点击按钮写入选项索引，关闭则写入 false。", true},
        {"交互", "显示变量", "显示变量 名称", "打印变量值。", true},
        {"交互", "帮助", "帮助", "列出命令目录。", true},
        {"交互", "列出命令", "列出命令", "列出全部命令及实现状态。", true},

        {"网络", "下载", "下载 网址 到 路径", "下载文件。", true},
        {"网络", "上传", "上传 路径 到 网址", "上传文件。", false},
        {"网络", "请求", "请求 GET|POST 网址 [数据] 到 变量名", "发送 HTTP 请求。", true},
        {"网络", "打开网页", "打开网页 网址", "打开默认浏览器。", true},
        {"网络", "读取网页", "读取网页 网址 到 变量名", "获取网页内容。", true},
        {"网络", "检测网络", "检测网络 到 变量名", "检查网络可达性。", true},
        {"网络", "本机IP", "本机IP 到 变量名", "读取本机 IP。", false},
        {"网络", "公网IP", "公网IP 到 变量名", "读取公网 IP。", false},
        {"网络", "端口检测", "端口检测 主机 端口 到 变量名", "检查端口可用性。", false},
        {"网络", "发送邮件", "发送邮件 配置 内容", "发送邮件。", false},

        {"时间", "当前时间", "当前时间 到 变量名", "读取当前时间。", true},
        {"时间", "当前日期", "当前日期 到 变量名", "读取当前日期。", true},
        {"时间", "格式时间", "格式时间 模板 到 变量名", "格式化当前时间。", true},
        {"时间", "定时", "定时 秒数 命令", "延时后执行命令。", true},
        {"时间", "每隔", "每隔 秒数 命令", "周期任务。", false},
        {"时间", "计时开始", "计时开始 名称", "开始计时。", false},
        {"时间", "计时结束", "计时结束 名称 到 变量名", "结束计时。", false},

        {"图像", "找图", "找图 图片路径 到 变量名", "屏幕找图。", true},
        {"图像", "识字", "识字 (x1,y1)->(x2,y2) 到 变量名", "OCR 识别。", true},
        {"图像", "取色", "取色 (x,y) 到 变量名", "读取像素颜色。", true},
        {"图像", "比色", "比色 (x,y) 颜色 到 变量名", "比较颜色。", true},
        {"图像", "找色", "找色 颜色 (x1,y1)->(x2,y2) 到 变量名", "区域找色，找到后返回坐标。", true},
        {"图像", "保存截图", "保存截图 路径", "保存当前屏幕截图。", true},

        {"调试", "日志", "日志 内容", "写调试日志。", true},
        {"调试", "警告", "警告 内容", "输出警告。", true},
        {"调试", "错误", "错误 内容", "输出错误。", true},
        {"调试", "断点", "断点", "在此处暂停。", true},
        {"调试", "追踪", "追踪 开|关", "打开或关闭命令追踪。", false},
        {"调试", "版本", "版本", "显示解释器版本。", true}
    };
}

class Interpreter {
public:
    Interpreter() : catalog_(buildCommandCatalog()) {}

    int runFile(const std::string& path) {
        const auto previousPath = currentScriptPath_;
        const auto previousLines = scriptLines_;
        const auto previousFunctions = functions_;

        currentScriptPath_ = std::filesystem::absolute(path).string();
        scriptLines_ = loadScriptLines(path);
        buildFunctionTable();
        executeRange(0, scriptLines_.size());

        currentScriptPath_ = previousPath;
        scriptLines_ = previousLines;
        functions_ = previousFunctions;
        return 0;
    }

private:
    std::map<std::string, std::string> variables_;
    std::vector<CommandInfo> catalog_;
    std::string currentScriptPath_;
    std::vector<ScriptLine> scriptLines_;
    std::map<std::string, FunctionInfo> functions_;
    std::vector<std::map<std::string, std::string>> localScopes_;
    bool returnTriggered_ = false;
    std::string returnValue_;
    int loopDepth_ = 0;

    std::string getVariableValue(const std::string& name) const {
        for (auto it = localScopes_.rbegin(); it != localScopes_.rend(); ++it) {
            auto found = it->find(name);
            if (found != it->end()) {
                return found->second;
            }
        }
        auto global = variables_.find(name);
        return global == variables_.end() ? "" : global->second;
    }

    bool hasVariable(const std::string& name) const {
        for (auto it = localScopes_.rbegin(); it != localScopes_.rend(); ++it) {
            if (it->find(name) != it->end()) {
                return true;
            }
        }
        return variables_.find(name) != variables_.end();
    }

    void setVariableValue(const std::string& name, const std::string& value) {
        if (!localScopes_.empty()) {
            localScopes_.back()[name] = value;
            return;
        }
        variables_[name] = value;
    }

    void eraseVariableValue(const std::string& name) {
        if (!localScopes_.empty()) {
            localScopes_.back().erase(name);
            return;
        }
        variables_.erase(name);
    }

    bool isCommentOrEmpty(const std::string& rawLine) const {
        const std::string line = trim(stripBom(rawLine));
        return line.empty() || startsWith(line, "#") || startsWith(line, "//") || startsWith(line, "注释 ");
    }

    std::string prepareLine(const std::string& rawLine) const {
        return expandVariables(trim(stripBom(rawLine)));
    }

    std::vector<ScriptLine> loadScriptLines(const std::string& path) const {
        std::ifstream input(path);
        if (!input) {
            throw std::runtime_error("无法打开脚本文件: " + path);
        }

        std::vector<ScriptLine> lines;
        std::string line;
        int lineNumber = 0;
        while (std::getline(input, line)) {
            ++lineNumber;
            lines.push_back({line, lineNumber});
        }
        return lines;
    }

    bool isFunctionCallSyntax(const std::string& line) const {
        std::string name;
        std::vector<std::string> args;
        if (!parseFunctionSignature(line, name, args)) {
            return false;
        }
        return functions_.find(name) != functions_.end();
    }

    bool isIfBlockStart(const std::vector<std::string>& tokens) const {
        return !tokens.empty() && tokens[0] == "如果" &&
               std::find(tokens.begin(), tokens.end(), "则") == tokens.end();
    }

    int indentLevel(const std::string& rawLine) const {
        int indent = 0;
        for (char ch : rawLine) {
            if (ch == ' ') {
                ++indent;
            } else if (ch == '\t') {
                indent += 4;
            } else {
                break;
            }
        }
        return indent;
    }

    size_t findMatchingEnd(size_t startIndex) const {
        int depth = 0;
        for (size_t i = startIndex + 1; i < scriptLines_.size(); ++i) {
            if (isCommentOrEmpty(scriptLines_[i].text)) {
                continue;
            }
            const auto tokens = tokenize(prepareLine(scriptLines_[i].text));
            if (tokens.empty()) {
                continue;
            }
            if (tokens[0] == "函数" || isIfBlockStart(tokens)) {
                ++depth;
            } else if (tokens[0] == "结束") {
                if (depth == 0) {
                    return i;
                }
                --depth;
            }
        }
        throw std::runtime_error("缺少匹配的 结束");
    }

    bool usesIndentedBlock(size_t startIndex) const {
        const int baseIndent = indentLevel(scriptLines_[startIndex].text);
        for (size_t i = startIndex + 1; i < scriptLines_.size(); ++i) {
            if (isCommentOrEmpty(scriptLines_[i].text)) {
                continue;
            }
            const auto tokens = tokenize(prepareLine(scriptLines_[i].text));
            if (!tokens.empty() && tokens[0] == "结束") {
                return false;
            }
            return indentLevel(scriptLines_[i].text) > baseIndent;
        }
        return false;
    }

    struct BlockBounds {
        size_t bodyStart = 0;
        size_t bodyEnd = 0;
        size_t nextIndex = 0;
    };

    BlockBounds getBlockBounds(size_t startIndex) const {
        BlockBounds bounds{};
        bounds.bodyStart = startIndex + 1;

        if (!usesIndentedBlock(startIndex)) {
            const size_t closeIndex = findMatchingEnd(startIndex);
            bounds.bodyEnd = closeIndex;
            bounds.nextIndex = closeIndex + 1;
            return bounds;
        }

        const int baseIndent = indentLevel(scriptLines_[startIndex].text);
        size_t i = startIndex + 1;
        while (i < scriptLines_.size()) {
            if (isCommentOrEmpty(scriptLines_[i].text)) {
                ++i;
                continue;
            }
            if (indentLevel(scriptLines_[i].text) <= baseIndent) {
                break;
            }
            ++i;
        }
        bounds.bodyEnd = i;
        bounds.nextIndex = i;
        return bounds;
    }

    void buildFunctionTable() {
        functions_.clear();
        for (size_t i = 0; i < scriptLines_.size(); ++i) {
            if (isCommentOrEmpty(scriptLines_[i].text)) {
                continue;
            }
            const auto tokens = tokenize(prepareLine(scriptLines_[i].text));
            if (tokens.empty() || tokens[0] != "函数") {
                continue;
            }
            if (tokens.size() < 2) {
                throw std::runtime_error("第 " + std::to_string(scriptLines_[i].lineNumber) + " 行函数定义缺少签名");
            }

            std::string name;
            std::vector<std::string> params;
            if (!parseFunctionSignature(join(tokens, 1), name, params)) {
                throw std::runtime_error("第 " + std::to_string(scriptLines_[i].lineNumber) + " 行函数签名格式错误");
            }

            FunctionInfo info;
            info.name = name;
            info.params = params;
            const BlockBounds bounds = getBlockBounds(i);
            info.startIndex = bounds.bodyStart;
            info.endIndex = bounds.bodyEnd;
            functions_[name] = info;
            i = bounds.nextIndex == 0 ? i : bounds.nextIndex - 1;
        }
    }

    std::map<std::string, size_t> buildLabelTable(size_t start, size_t end) const {
        std::map<std::string, size_t> labels;
        for (size_t i = start; i < end; ++i) {
            if (isCommentOrEmpty(scriptLines_[i].text)) {
                continue;
            }
            const auto tokens = tokenize(prepareLine(scriptLines_[i].text));
            if (tokens.empty()) {
                continue;
            }
            if (tokens[0] == "函数") {
                const BlockBounds bounds = getBlockBounds(i);
                i = bounds.nextIndex == 0 ? i : bounds.nextIndex - 1;
                continue;
            }
            if (tokens[0] == "标签" && tokens.size() >= 2) {
                labels[tokens[1]] = i;
            }
        }
        return labels;
    }

    std::string expandVariables(std::string text) const {
        size_t pos = 0;
        while ((pos = text.find("${", pos)) != std::string::npos) {
            const size_t end = text.find('}', pos + 2);
            if (end == std::string::npos) {
                break;
            }
            const std::string key = text.substr(pos + 2, end - pos - 2);
            const std::string replacement = getVariableValue(key);
            text.replace(pos, end - pos + 1, replacement);
            pos += replacement.size();
        }
        return text;
    }

    bool compareValues(const std::string& left, const std::string& op, const std::string& right) const {
        const bool maybeNumber =
            !left.empty() && !right.empty() &&
            (std::isdigit(static_cast<unsigned char>(left[0])) || left[0] == '-' || left[0] == '+') &&
            (std::isdigit(static_cast<unsigned char>(right[0])) || right[0] == '-' || right[0] == '+');

        if (maybeNumber) {
            const double a = parseNumber(left);
            const double b = parseNumber(right);
            if (op == "=" || op == "==") return a == b;
            if (op == "!=") return a != b;
            if (op == ">") return a > b;
            if (op == "<") return a < b;
            if (op == ">=") return a >= b;
            if (op == "<=") return a <= b;
        } else {
            if (op == "=" || op == "==") return left == right;
            if (op == "!=") return left != right;
            if (op == "包含") return left.find(right) != std::string::npos;
            if (op == "不含") return left.find(right) == std::string::npos;
            if (op == ">") return left > right;
            if (op == "<") return left < right;
            if (op == ">=") return left >= right;
            if (op == "<=") return left <= right;
        }

        throw std::runtime_error("不支持的比较运算符: " + op);
    }

    bool evaluateConditionTokens(const std::vector<std::string>& tokens, size_t start) const {
        if (tokens.size() < start + 3) {
            throw std::runtime_error("条件表达式至少需要 左值 运算符 右值");
        }
        return compareValues(tokens[start], tokens[start + 1], tokens[start + 2]);
    }

    void commandAssign(const std::vector<std::string>& tokens) {
        if (tokens.size() < 3 || tokens[1] != "=") {
            throw std::runtime_error("赋值语法: 变量名 = 值");
        }
        const std::string& target = tokens[0];
        if (tokens.size() == 3) {
            std::string functionName;
            std::vector<std::string> args;
            if (parseFunctionSignature(tokens[2], functionName, args) && functions_.find(functionName) != functions_.end()) {
                setVariableValue(target, callFunctionByName(functionName, args));
                return;
            }
            setVariableValue(target, tokens[2]);
            return;
        }
        if (tokens.size() >= 5) {
            const std::string left = tokens[2];
            const std::string op = tokens[3];
            const std::string right = join(tokens, 4);
            if (op == "加") {
                setVariableValue(target, toDisplayNumber(parseNumber(left) + parseNumber(right)));
                return;
            }
            if (op == "减") {
                setVariableValue(target, toDisplayNumber(parseNumber(left) - parseNumber(right)));
                return;
            }
            if (op == "乘") {
                setVariableValue(target, toDisplayNumber(parseNumber(left) * parseNumber(right)));
                return;
            }
            if (op == "除") {
                setVariableValue(target, toDisplayNumber(parseNumber(left) / parseNumber(right)));
                return;
            }
            if (op == "连接") {
                setVariableValue(target, left + right);
                return;
            }
        }
        setVariableValue(target, join(tokens, 2));
    }

    std::string callFunctionByName(const std::string& name, const std::vector<std::string>& args) {
        auto it = functions_.find(name);
        if (it == functions_.end()) {
            throw std::runtime_error("函数不存在: " + name);
        }
        const FunctionInfo& function = it->second;
        if (args.size() != function.params.size()) {
            throw std::runtime_error("函数参数数量不匹配: " + name);
        }

        localScopes_.push_back({});
        for (size_t i = 0; i < function.params.size(); ++i) {
            localScopes_.back()[function.params[i]] = args[i];
        }

        const bool previousReturn = returnTriggered_;
        const std::string previousValue = returnValue_;
        returnTriggered_ = false;
        returnValue_.clear();
        executeRange(function.startIndex, function.endIndex);
        const std::string value = returnValue_;
        returnTriggered_ = previousReturn;
        returnValue_ = previousValue;
        localScopes_.pop_back();
        return value;
    }

    size_t executeIfBlock(size_t startIndex, size_t endIndex) {
        if (!usesIndentedBlock(startIndex)) {
            std::vector<size_t> branchLines;
            branchLines.push_back(startIndex);
            int depth = 0;
            size_t blockEnd = static_cast<size_t>(-1);

            for (size_t i = startIndex + 1; i < endIndex; ++i) {
                if (isCommentOrEmpty(scriptLines_[i].text)) {
                    continue;
                }
                const auto tokens = tokenize(prepareLine(scriptLines_[i].text));
                if (tokens.empty()) {
                    continue;
                }
                if (tokens[0] == "函数" || isIfBlockStart(tokens)) {
                    ++depth;
                    continue;
                }
                if (tokens[0] == "结束") {
                    if (depth == 0) {
                        blockEnd = i;
                        break;
                    }
                    --depth;
                    continue;
                }
                if (depth == 0 && (tokens[0] == "否则如果" || tokens[0] == "否则")) {
                    branchLines.push_back(i);
                }
            }

            if (blockEnd == static_cast<size_t>(-1)) {
                throw std::runtime_error("如果 代码块缺少 结束");
            }

            bool executed = false;
            for (size_t b = 0; b < branchLines.size(); ++b) {
                const size_t lineIndex = branchLines[b];
                const auto tokens = tokenize(prepareLine(scriptLines_[lineIndex].text));
                bool run = false;
                if (tokens[0] == "如果" || tokens[0] == "否则如果") {
                    run = evaluateConditionTokens(tokens, 1);
                } else if (tokens[0] == "否则") {
                    run = true;
                }

                const size_t bodyStart = lineIndex + 1;
                const size_t bodyEnd = (b + 1 < branchLines.size()) ? branchLines[b + 1] : blockEnd;
                if (!executed && run) {
                    executeRange(bodyStart, bodyEnd);
                    executed = true;
                }
            }

            return blockEnd + 1;
        }

        std::vector<size_t> branchLines;
        branchLines.push_back(startIndex);
        const int baseIndent = indentLevel(scriptLines_[startIndex].text);
        size_t chainEnd = endIndex;

        for (size_t i = startIndex + 1; i < endIndex; ++i) {
            if (isCommentOrEmpty(scriptLines_[i].text)) {
                continue;
            }
            const auto tokens = tokenize(prepareLine(scriptLines_[i].text));
            if (tokens.empty()) {
                continue;
            }
            const int currentIndent = indentLevel(scriptLines_[i].text);
            if (currentIndent < baseIndent) {
                chainEnd = i;
                break;
            }
            if (currentIndent == baseIndent &&
                tokens[0] != "否则如果" &&
                tokens[0] != "否则") {
                chainEnd = i;
                break;
            }

            if (currentIndent == baseIndent &&
                (tokens[0] == "否则如果" || tokens[0] == "否则")) {
                branchLines.push_back(i);
            }
        }

        bool executed = false;
        for (size_t b = 0; b < branchLines.size(); ++b) {
            const size_t lineIndex = branchLines[b];
            const auto tokens = tokenize(prepareLine(scriptLines_[lineIndex].text));
            bool run = false;
            if (tokens[0] == "如果" || tokens[0] == "否则如果") {
                run = evaluateConditionTokens(tokens, 1);
            } else if (tokens[0] == "否则") {
                run = true;
            }

            const size_t bodyStart = lineIndex + 1;
            const size_t bodyEnd = (b + 1 < branchLines.size()) ? branchLines[b + 1] : chainEnd;
            if (!executed && run) {
                executeRange(bodyStart, bodyEnd);
                executed = true;
            }
        }

        return chainEnd;
    }

    size_t executeRepeatBlock(size_t startIndex, size_t endIndex) {
        const auto headerTokens = tokenize(prepareLine(scriptLines_[startIndex].text));
        if (headerTokens.size() < 2) {
            throw std::runtime_error("重复 语法: 重复 次数 [命令]");
        }
        const int count = static_cast<int>(parseNumber(headerTokens[1]));

        if (headerTokens.size() > 2) {
            const std::string nested = join(headerTokens, 2);
            ++loopDepth_;
            for (int i = 0; i < count; ++i) {
                try {
                    executeLine(nested, scriptLines_[startIndex].lineNumber);
                } catch (const ContinueSignal&) {
                    continue;
                } catch (const BreakSignal&) {
                    break;
                }
            }
            --loopDepth_;
            return startIndex + 1;
        }

        const BlockBounds bounds = getBlockBounds(startIndex);
        ++loopDepth_;
        for (int i = 0; i < count; ++i) {
            try {
                executeRange(bounds.bodyStart, bounds.bodyEnd);
            } catch (const ContinueSignal&) {
                continue;
            } catch (const BreakSignal&) {
                break;
            }
        }
        --loopDepth_;
        return bounds.nextIndex;
    }

    size_t executeWhileBlock(size_t startIndex, size_t endIndex) {
        const auto headerTokens = tokenize(prepareLine(scriptLines_[startIndex].text));
        if (headerTokens.size() < 4) {
            throw std::runtime_error("当 语法: 当 左值 运算符 右值 [时 命令]");
        }

        if (headerTokens.size() > 4) {
            if (headerTokens[4] != "时") {
                throw std::runtime_error("当 单行语法需要使用 时");
            }
            const std::string nested = join(headerTokens, 5);
            ++loopDepth_;
            while (evaluateConditionTokens(headerTokens, 1)) {
                try {
                    executeLine(nested, scriptLines_[startIndex].lineNumber);
                } catch (const ContinueSignal&) {
                    continue;
                } catch (const BreakSignal&) {
                    break;
                }
            }
            --loopDepth_;
            return startIndex + 1;
        }

        const BlockBounds bounds = getBlockBounds(startIndex);
        ++loopDepth_;
        while (evaluateConditionTokens(headerTokens, 1)) {
            try {
                executeRange(bounds.bodyStart, bounds.bodyEnd);
            } catch (const ContinueSignal&) {
                continue;
            } catch (const BreakSignal&) {
                break;
            }
        }
        --loopDepth_;
        return bounds.nextIndex;
    }

    size_t executeForeachBlock(size_t startIndex, size_t endIndex) {
        const auto headerTokens = tokenize(prepareLine(scriptLines_[startIndex].text));
        if (headerTokens.size() < 4 || headerTokens[2] != "在") {
            throw std::runtime_error("遍历 语法: 遍历 变量 在 列表 [命令]");
        }
        const std::string variableName = headerTokens[1];
        const std::vector<std::string> items = splitListItems(headerTokens[3]);

        if (headerTokens.size() > 4) {
            const std::string nested = join(headerTokens, 4);
            ++loopDepth_;
            for (const auto& item : items) {
                setVariableValue(variableName, item);
                try {
                    executeLine(nested, scriptLines_[startIndex].lineNumber);
                } catch (const ContinueSignal&) {
                    continue;
                } catch (const BreakSignal&) {
                    break;
                }
            }
            --loopDepth_;
            return startIndex + 1;
        }

        const BlockBounds bounds = getBlockBounds(startIndex);
        ++loopDepth_;
        for (const auto& item : items) {
            setVariableValue(variableName, item);
            try {
                executeRange(bounds.bodyStart, bounds.bodyEnd);
            } catch (const ContinueSignal&) {
                continue;
            } catch (const BreakSignal&) {
                break;
            }
        }
        --loopDepth_;
        return bounds.nextIndex;
    }

    void executeRange(size_t start, size_t end) {
        auto labels = buildLabelTable(start, end);
        size_t index = start;
        while (index < end) {
            if (returnTriggered_) {
                return;
            }
            const ScriptLine& scriptLine = scriptLines_[index];
            if (isCommentOrEmpty(scriptLine.text)) {
                ++index;
                continue;
            }

            const std::string line = prepareLine(scriptLine.text);
            const auto tokens = tokenize(line);
            if (tokens.empty()) {
                ++index;
                continue;
            }

            if (tokens[0] == "函数") {
                index = getBlockBounds(index).nextIndex;
                continue;
            }
            if (tokens[0] == "标签") {
                ++index;
                continue;
            }
            if (tokens[0] == "结束") {
                return;
            }
            if (tokens[0] == "如果" && std::find(tokens.begin(), tokens.end(), "则") == tokens.end()) {
                index = executeIfBlock(index, end);
                continue;
            }
            if (tokens[0] == "当") {
                index = executeWhileBlock(index, end);
                continue;
            }
            if (tokens[0] == "重复") {
                index = executeRepeatBlock(index, end);
                continue;
            }
            if (tokens[0] == "遍历") {
                index = executeForeachBlock(index, end);
                continue;
            }
            if (tokens[0] == "否则如果" || tokens[0] == "否则") {
                return;
            }
            if (tokens[0] == "去") {
                if (tokens.size() < 2) {
                    throw std::runtime_error("第 " + std::to_string(scriptLine.lineNumber) + " 行执行失败: 去 需要标签名");
                }
                auto it = labels.find(tokens[1]);
                if (it == labels.end()) {
                    throw std::runtime_error("第 " + std::to_string(scriptLine.lineNumber) + " 行执行失败: 标签不存在: " + tokens[1]);
                }
                index = it->second + 1;
                continue;
            }

            executeLine(line, scriptLine.lineNumber);
            ++index;
        }
    }

    void executeLine(const std::string& rawLine, int lineNumber) {
        std::string line = trim(stripBom(rawLine));
        if (line.empty() || startsWith(line, "#") || startsWith(line, "//")) {
            return;
        }
        if (startsWith(line, "注释 ")) {
            return;
        }
        line = expandVariables(line);
        const auto tokens = tokenize(line);
        if (tokens.empty()) {
            return;
        }

        try {
            const std::string& command = tokens[0];
            if (tokens.size() >= 3 && tokens[1] == "=" && command != "变量") {
                commandAssign(tokens);
            } else if (command == "如果") {
                if (tokens.size() < 6) {
                    throw std::runtime_error("如果 语法: 如果 左值 运算符 右值 则 命令");
                }
                if (tokens[4] != "则") {
                    throw std::runtime_error("如果 当前仅支持单行语法，并且必须包含 则");
                }
                if (compareValues(tokens[1], tokens[2], tokens[3])) {
                    executeLine(join(tokens, 5), lineNumber);
                }
            } else if (command == "重复") {
                if (tokens.size() < 3) {
                    return;
                }
                const int count = static_cast<int>(parseNumber(tokens[1]));
                const std::string nested = join(tokens, 2);
                ++loopDepth_;
                for (int i = 0; i < count; ++i) {
                    try {
                        executeLine(nested, lineNumber);
                    } catch (const ContinueSignal&) {
                        continue;
                    } catch (const BreakSignal&) {
                        break;
                    }
                }
                --loopDepth_;
            } else if (command == "当") {
                if (tokens.size() < 6 || tokens[4] != "时") {
                    return;
                }
                ++loopDepth_;
                while (compareValues(tokens[1], tokens[2], tokens[3])) {
                    try {
                        executeLine(join(tokens, 5), lineNumber);
                    } catch (const ContinueSignal&) {
                        continue;
                    } catch (const BreakSignal&) {
                        break;
                    }
                }
                --loopDepth_;
            } else if (command == "遍历") {
                if (tokens.size() < 5 || tokens[2] != "在") {
                    return;
                }
                const auto items = splitListItems(tokens[3]);
                const std::string nested = join(tokens, 4);
                ++loopDepth_;
                for (const auto& item : items) {
                    setVariableValue(tokens[1], item);
                    try {
                        executeLine(nested, lineNumber);
                    } catch (const ContinueSignal&) {
                        continue;
                    } catch (const BreakSignal&) {
                        break;
                    }
                }
                --loopDepth_;
            } else if (command == "函数" || command == "标签" || command == "去" ||
                       command == "否则如果" || command == "否则" || command == "结束") {
                return;
            } else if (command == "跳出") {
                if (loopDepth_ <= 0) {
                    throw std::runtime_error("跳出 只能在循环中使用");
                }
                throw BreakSignal{};
            } else if (command == "继续") {
                if (loopDepth_ <= 0) {
                    throw std::runtime_error("继续 只能在循环中使用");
                }
                throw ContinueSignal{};
            } else if (command == "点击") {
                commandClick(tokens);
            } else if (command == "移动鼠标") {
                commandMoveMouse(tokens);
            } else if (command == "滑动") {
                commandSwipe(tokens);
            } else if (command == "滚轮") {
                commandWheel(tokens);
            } else if (command == "拖到") {
                commandDrag(tokens);
            } else if (command == "鼠标回中") {
                SetCursorPos(GetSystemMetrics(SM_CXSCREEN) / 2, GetSystemMetrics(SM_CYSCREEN) / 2);
            } else if (command == "按键") {
                commandKey(tokens);
            } else if (command == "回车") {
                sendVirtualKey(VK_RETURN);
            } else if (command == "退格") {
                sendVirtualKey(VK_BACK);
            } else if (command == "删除键") {
                sendVirtualKey(VK_DELETE);
            } else if (command == "执行") {
                commandExecute(tokens);
            } else if (command == "打开" || command == "打开网页") {
                commandOpen(tokens);
            } else if (command == "激活窗口") {
                commandActivateWindow(tokens);
            } else if (command == "关闭窗口") {
                commandWindowAction(tokens, WM_CLOSE, SW_SHOWNORMAL);
            } else if (command == "最小化窗口") {
                commandWindowAction(tokens, 0, SW_MINIMIZE);
            } else if (command == "最大化窗口") {
                commandWindowAction(tokens, 0, SW_MAXIMIZE);
            } else if (command == "还原窗口") {
                commandWindowAction(tokens, 0, SW_RESTORE);
            } else if (command == "等待") {
                commandWait(tokens);
            } else if (command == "定时") {
                if (tokens.size() < 3) {
                    throw std::runtime_error("定时 语法: 定时 秒数 命令");
                }
                commandWait({"等待", tokens[1]});
                executeLine(join(tokens, 2), lineNumber);
            } else if (command == "暂停") {
                std::cout << "按回车继续..." << std::endl;
                std::string ignored;
                std::getline(std::cin, ignored);
            } else if (command == "退出脚本") {
                const int exitCode = tokens.size() >= 2 ? static_cast<int>(parseNumber(tokens[1])) : 0;
                std::exit(exitCode);
            } else if (command == "蜂鸣") {
                if (tokens.size() < 3) {
                    throw std::runtime_error("蜂鸣 需要频率和毫秒");
                }
                Beep(static_cast<DWORD>(parseNumber(tokens[1])), static_cast<DWORD>(parseNumber(tokens[2])));
            } else if (command == "输出") {
                std::cout << join(tokens, 1) << std::endl;
            } else if (command == "换行") {
                std::cout << std::endl;
            } else if (command == "清屏") {
                system("cls");
            } else if (command == "标题") {
                SetConsoleTitleW(utf8ToWide(join(tokens, 1)).c_str());
            } else if (command == "变量") {
                commandSet(tokens);
            } else if (command == "加" || command == "减" || command == "乘" || command == "除") {
                commandMath(tokens);
            } else if (command == "连接") {
                commandConcat(tokens);
            } else if (command == "删除变量") {
                if (tokens.size() < 2) {
                    throw std::runtime_error("删除变量 需要变量名");
                }
                eraseVariableValue(tokens[1]);
            } else if (command == "读取环境") {
                commandReadEnv(tokens);
            } else if (command == "设置环境") {
                commandSetEnv(tokens);
            } else if (command == "转整数" || command == "转小数" || command == "转文本") {
                commandCast(tokens);
            } else if (command == "输入文本") {
                commandInputText(tokens);
            } else if (command == "询问") {
                commandAsk(tokens);
            } else if (command == "确认") {
                commandConfirm(tokens);
            } else if (command == "确定弹窗") {
                commandAlertPopup(tokens);
            } else if (command == "询问弹窗") {
                commandInputPopup(tokens);
            } else if (command == "选择弹窗") {
                commandSelectPopup(tokens);
            } else if (command == "当前时间") {
                commandTime(tokens, "%H:%M:%S");
            } else if (command == "当前日期") {
                commandTime(tokens, "%Y-%m-%d");
            } else if (command == "格式时间") {
                commandFormatTime(tokens);
            } else if (command == "显示变量") {
                if (tokens.size() < 2) {
                    throw std::runtime_error("显示变量 需要变量名");
                }
                std::cout << tokens[1] << " = " << getVariableValue(tokens[1]) << std::endl;
            } else if (command == "复制文件") {
                commandCopyFile(tokens);
            } else if (command == "移动文件") {
                commandMoveFile(tokens);
            } else if (command == "删除文件") {
                if (tokens.size() < 2) {
                    throw std::runtime_error("删除文件 需要路径");
                }
                std::filesystem::remove(tokens[1]);
            } else if (command == "创建文件") {
                if (tokens.size() < 2) {
                    throw std::runtime_error("创建文件 需要路径");
                }
                std::ofstream output(tokens[1], std::ios::binary);
            } else if (command == "写入文件") {
                commandWriteFile(tokens, false);
            } else if (command == "追加文件") {
                commandWriteFile(tokens, true);
            } else if (command == "读取文件") {
                commandReadFile(tokens);
            } else if (command == "创建目录") {
                if (tokens.size() < 2) {
                    throw std::runtime_error("创建目录 需要路径");
                }
                std::filesystem::create_directories(tokens[1]);
            } else if (command == "删除目录") {
                if (tokens.size() < 2) {
                    throw std::runtime_error("删除目录 需要路径");
                }
                std::filesystem::remove_all(tokens[1]);
            } else if (command == "列出目录") {
                commandListDir(tokens);
            } else if (command == "文件存在") {
                commandExists(tokens, false);
            } else if (command == "路径存在") {
                commandExists(tokens, true);
            } else if (command == "文件大小") {
                commandFileSize(tokens);
            } else if (command == "改名文件") {
                if (tokens.size() < 3) {
                    throw std::runtime_error("改名文件 需要旧路径和新路径");
                }
                std::filesystem::rename(tokens[1], tokens[2]);
            } else if (command == "调用") {
                if (tokens.size() < 2) {
                    throw std::runtime_error("调用 需要脚本路径");
                }
                std::filesystem::path base = std::filesystem::path(currentScriptPath_).parent_path();
                std::filesystem::path child = std::filesystem::path(tokens[1]).is_absolute()
                    ? std::filesystem::path(tokens[1])
                    : (base / tokens[1]);
                runFile(child.string());
            } else if (command == "下载") {
                commandDownload(tokens);
            } else if (command == "读取网页") {
                commandReadWeb(tokens);
            } else if (command == "请求") {
                commandRequest(tokens);
            } else if (command == "检测网络") {
                commandDetectNetwork(tokens);
            } else if (command == "找图") {
                commandFindImage(tokens);
            } else if (command == "识字") {
                commandOcr(tokens);
            } else if (command == "取色") {
                commandPickColor(tokens);
            } else if (command == "比色") {
                commandCompareColor(tokens);
            } else if (command == "找色") {
                commandFindColor(tokens);
            } else if (command == "保存截图") {
                commandSaveScreenshot(tokens);
            } else if (command == "返回") {
                returnTriggered_ = true;
                returnValue_ = tokens.size() >= 2 ? join(tokens, 1) : "";
            } else if (isFunctionCallSyntax(line)) {
                std::string functionName;
                std::vector<std::string> args;
                parseFunctionSignature(line, functionName, args);
                callFunctionByName(functionName, args);
            } else if (command == "日志") {
                std::cout << "[日志] " << join(tokens, 1) << std::endl;
            } else if (command == "警告") {
                std::cout << "[警告] " << join(tokens, 1) << std::endl;
            } else if (command == "错误") {
                std::cerr << "[错误] " << join(tokens, 1) << std::endl;
            } else if (command == "断点") {
                std::cout << "[断点] 按回车继续..." << std::endl;
                std::string ignored;
                std::getline(std::cin, ignored);
            } else if (command == "帮助" || command == "列出命令") {
                commandHelp();
            } else if (command == "版本") {
                std::cout << "OperationScript explain.cpp 0.4.0" << std::endl;
            } else {
                throw std::runtime_error("未实现或未知命令: " + command);
            }
        } catch (const std::exception& ex) {
            std::ostringstream out;
            out << "第 " << lineNumber << " 行执行失败: " << ex.what();
            throw std::runtime_error(out.str());
        }
    }

    void commandClick(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4) {
            throw std::runtime_error("点击 语法不足");
        }
        const bool left = tokens[1] == "左键";
        if (!left && tokens[1] != "右键") {
            throw std::runtime_error("点击 键型只能是 左键 或 右键");
        }
        const std::string type = tokens[2];
        const Point point = parsePoint(tokens[3]);
        SetCursorPos(point.x, point.y);
        sleepMs(40);

        if (type == "单点") {
            sendMouseClick(left);
            return;
        }
        if (type == "双击") {
            sendMouseClick(left);
            sleepMs(60);
            sendMouseClick(left);
            return;
        }
        if (type == "多击") {
            if (tokens.size() < 5) {
                throw std::runtime_error("多击 需要次数");
            }
            const int count = static_cast<int>(parseNumber(tokens[4]));
            for (int i = 0; i < count; ++i) {
                sendMouseClick(left);
                sleepMs(60);
            }
            return;
        }
        if (type == "长按") {
            if (tokens.size() < 5) {
                throw std::runtime_error("长按 需要秒数");
            }
            const double seconds = parseNumber(tokens[4]);
            sendMouseDown(left);
            sleepSeconds(seconds);
            sendMouseUp(left);
            return;
        }
        throw std::runtime_error("点击 类型只能是 单点 / 双击 / 多击 / 长按");
    }

    void commandMoveMouse(const std::vector<std::string>& tokens) {
        if (tokens.size() < 2) {
            throw std::runtime_error("移动鼠标 需要坐标");
        }
        const Point point = parsePoint(tokens[1]);
        SetCursorPos(point.x, point.y);
    }

    void commandWheel(const std::vector<std::string>& tokens) {
        if (tokens.size() < 3) {
            throw std::runtime_error("滚轮 需要方向和次数");
        }
        const int count = static_cast<int>(parseNumber(tokens[2]));
        const int delta = tokens[1] == "上" ? WHEEL_DELTA : -WHEEL_DELTA;
        for (int i = 0; i < std::abs(count); ++i) {
            sendMouseWheel(delta);
            sleepMs(25);
        }
    }

    void commandDrag(const std::vector<std::string>& tokens) {
        if (tokens.size() < 3) {
            throw std::runtime_error("拖到 语法: 拖到 左键|右键 (x1,y1)->(x2,y2)");
        }
        std::vector<std::string> nested = {"滑动", tokens[1], "0.5", "线性", tokens[2]};
        commandSwipe(nested);
    }

    void commandKey(const std::vector<std::string>& tokens) {
        if (tokens.size() < 2) {
            throw std::runtime_error("按键 需要键名");
        }

        static const std::map<std::string, WORD> keyMap = {
            {"回车", VK_RETURN}, {"退格", VK_BACK}, {"删除", VK_DELETE},
            {"空格", VK_SPACE}, {"Tab", VK_TAB}, {"ESC", VK_ESCAPE},
            {"上", VK_UP}, {"下", VK_DOWN}, {"左", VK_LEFT}, {"右", VK_RIGHT}
        };

        auto it = keyMap.find(tokens[1]);
        if (it != keyMap.end()) {
            sendVirtualKey(it->second);
            return;
        }

        if (tokens[1].size() == 1) {
            SHORT value = VkKeyScanA(tokens[1][0]);
            if (value == -1) {
                throw std::runtime_error("无法识别按键: " + tokens[1]);
            }
            sendVirtualKey(static_cast<WORD>(value & 0xff));
            return;
        }

        throw std::runtime_error("无法识别按键: " + tokens[1]);
    }

    void commandSwipe(const std::vector<std::string>& tokens) {
        if (tokens.size() < 5) {
            throw std::runtime_error("滑动 语法不足");
        }
        const bool left = tokens[1] == "左键";
        if (!left && tokens[1] != "右键") {
            throw std::runtime_error("滑动 键型只能是 左键 或 右键");
        }
        const double seconds = parseNumber(tokens[2]);
        const std::string easing = tokens[3];
        const std::string& segment = tokens[4];
        const auto arrow = segment.find("->");
        if (arrow == std::string::npos) {
            throw std::runtime_error("滑动 轨迹格式应为 (x1,y1)->(x2,y2)");
        }
        const Point from = parsePoint(segment.substr(0, arrow));
        const Point to = parsePoint(segment.substr(arrow + 2));

        const int steps = std::max(10, static_cast<int>(seconds * 120.0));
        SetCursorPos(from.x, from.y);
        sleepMs(20);
        sendMouseDown(left);
        for (int i = 0; i <= steps; ++i) {
            const double t = static_cast<double>(i) / static_cast<double>(steps);
            const double p = easeProgress(t, easing);
            const int x = static_cast<int>(std::round(from.x + (to.x - from.x) * p));
            const int y = static_cast<int>(std::round(from.y + (to.y - from.y) * p));
            SetCursorPos(x, y);
            sleepMs(std::max(1, static_cast<int>(seconds * 1000.0 / steps)));
        }
        sendMouseUp(left);
    }

    void commandExecute(const std::vector<std::string>& tokens) {
        if (tokens.size() < 3) {
            throw std::runtime_error("执行 需要 shellid 和命令");
        }
        const int shellId = static_cast<int>(parseNumber(tokens[1]));
        const std::string command = join(tokens, 2);
        std::string finalCommand;
        if (shellId == 0) {
            finalCommand = "cmd /c \"" + command + "\"";
        } else if (shellId == 1) {
            finalCommand = "powershell -Command \"" + command + "\"";
        } else {
            throw std::runtime_error("执行 shellid 只能是 0 或 1");
        }
        const int code = system(finalCommand.c_str());
        if (code != 0) {
            std::ostringstream out;
            out << "执行 返回代码: " << code;
            throw std::runtime_error(out.str());
        }
    }

    void commandOpen(const std::vector<std::string>& tokens) {
        if (tokens.size() < 2) {
            throw std::runtime_error("打开 需要目标");
        }
        const std::wstring target = utf8ToWide(join(tokens, 1));
        HINSTANCE result = ShellExecuteW(nullptr, L"open", target.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
        if (reinterpret_cast<INT_PTR>(result) <= 32) {
            throw std::runtime_error("打开 失败");
        }
    }

    void commandWait(const std::vector<std::string>& tokens) {
        if (tokens.size() < 2) {
            throw std::runtime_error("等待 需要秒数");
        }
        const double seconds = parseNumber(tokens[1]);
        sleepSeconds(seconds);
    }

    void commandSet(const std::vector<std::string>& tokens) {
        if (tokens.size() < 2) {
            throw std::runtime_error("变量 语法: 变量 名称 [= 值]");
        }
        if (tokens.size() == 2) {
            setVariableValue(tokens[1], "0");
            return;
        }
        if (tokens.size() >= 4 && tokens[2] == "=") {
            setVariableValue(tokens[1], join(tokens, 3));
            return;
        }
        throw std::runtime_error("变量 语法: 变量 名称 [= 值]");
    }

    void commandMath(const std::vector<std::string>& tokens) {
        if (tokens.size() < 3) {
            throw std::runtime_error("算术命令需要变量名和数值");
        }
        double current = 0.0;
        const std::string currentValue = getVariableValue(tokens[1]);
        if (!currentValue.empty()) {
            current = parseNumber(currentValue);
        }
        const double value = parseNumber(tokens[2]);
        if (tokens[0] == "加") {
            current += value;
        } else if (tokens[0] == "减") {
            current -= value;
        } else if (tokens[0] == "乘") {
            current *= value;
        } else if (tokens[0] == "除") {
            current /= value;
        }
        setVariableValue(tokens[1], toDisplayNumber(current));
    }

    void commandConcat(const std::vector<std::string>& tokens) {
        if (tokens.size() < 3) {
            throw std::runtime_error("连接 需要变量名和值");
        }
        setVariableValue(tokens[1], getVariableValue(tokens[1]) + join(tokens, 2));
    }

    void commandReadEnv(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4 || tokens[2] != "到") {
            throw std::runtime_error("读取环境 语法: 读取环境 变量名 到 名称");
        }
        const char* value = std::getenv(tokens[1].c_str());
        setVariableValue(tokens[3], value ? value : "");
    }

    void commandSetEnv(const std::vector<std::string>& tokens) {
        if (tokens.size() < 3) {
            throw std::runtime_error("设置环境 需要变量名和值");
        }
        SetEnvironmentVariableA(tokens[1].c_str(), join(tokens, 2).c_str());
    }

    void commandCast(const std::vector<std::string>& tokens) {
        if (tokens.size() < 2) {
            throw std::runtime_error("类型转换需要变量名");
        }
        if (!hasVariable(tokens[1])) {
            setVariableValue(tokens[1], "");
            return;
        }
        std::string value = getVariableValue(tokens[1]);

        if (tokens[0] == "转整数") {
            setVariableValue(tokens[1], std::to_string(static_cast<int>(parseNumber(value))));
        } else if (tokens[0] == "转小数") {
            setVariableValue(tokens[1], toDisplayNumber(parseNumber(value)));
        }
    }

    void commandInputText(const std::vector<std::string>& tokens) {
        if (tokens.size() < 2) {
            throw std::runtime_error("输入文本 需要内容");
        }
        const std::wstring text = utf8ToWide(join(tokens, 1));
        for (wchar_t ch : text) {
            sendUnicodeChar(ch);
            sleepMs(8);
        }
    }

    void commandAsk(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4 || tokens[tokens.size() - 2] != "到") {
            throw std::runtime_error("询问 语法: 询问 提示 到 变量名");
        }
        std::cout << join(tokens, 1) << std::endl;
        std::string value;
        std::getline(std::cin, value);
        setVariableValue(tokens.back(), value);
    }

    void commandConfirm(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4 || tokens[tokens.size() - 2] != "到") {
            throw std::runtime_error("确认 语法: 确认 提示 到 变量名");
        }
        std::cout << join(tokens, 1) << " (y/n)" << std::endl;
        std::string value;
        std::getline(std::cin, value);
        setVariableValue(tokens.back(), (value == "y" || value == "Y" || value == "是") ? "真" : "假");
    }

    void commandAlertPopup(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4) {
            throw std::runtime_error("确定弹窗 语法: 确定弹窗 (x,y) 消息 变量名[OK文字]");
        }
        const auto target = splitNameAndBracket(tokens[3]);
        PopupDialogState state{};
        state.origin = parsePoint(tokens[1]);
        state.message = utf8ToWide(tokens[2]);
        state.buttonText = utf8ToWide(target.second.empty() ? "确定" : target.second);
        setVariableValue(target.first, runPopupDialog(state, L"确定弹窗") ? "true" : "false");
    }

    void commandInputPopup(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4) {
            throw std::runtime_error("询问弹窗 语法: 询问弹窗 (x,y) 消息 变量名[OK文字]");
        }
        const auto target = splitNameAndBracket(tokens[3]);

        PopupDialogState state{};
        state.origin = parsePoint(tokens[1]);
        state.message = utf8ToWide(tokens[2]);
        state.buttonText = utf8ToWide(target.second.empty() ? "确定" : target.second);
        state.showEdit = true;
        if (runPopupDialog(state, L"询问弹窗")) {
            setVariableValue(target.first, wideToUtf8(state.inputValue));
        } else {
            setVariableValue(target.first, "false");
        }
    }

    void commandSelectPopup(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4) {
            throw std::runtime_error("选择弹窗 语法: 选择弹窗 (x,y) 消息 变量名[选项1,选项2,...]");
        }
        const auto target = splitNameAndBracket(tokens[3]);
        if (target.second.empty()) {
            throw std::runtime_error("选择弹窗 至少需要一个选项");
        }

        PopupDialogState state{};
        state.origin = parsePoint(tokens[1]);
        state.message = utf8ToWide(tokens[2]);
        state.showSelect = true;
        for (const auto& option : parseOptionList(target.second)) {
            state.options.push_back(utf8ToWide(option));
        }
        if (runPopupDialog(state, L"选择弹窗")) {
            setVariableValue(target.first, std::to_string(state.selectedIndex));
        } else {
            setVariableValue(target.first, "false");
        }
    }

    void commandActivateWindow(const std::vector<std::string>& tokens) {
        if (tokens.size() < 2) {
            throw std::runtime_error("激活窗口 需要标题");
        }
        HWND hwnd = findWindowByTitle(utf8ToWide(join(tokens, 1)));
        if (hwnd == nullptr) {
            throw std::runtime_error("未找到窗口: " + join(tokens, 1));
        }
        ShowWindow(hwnd, SW_RESTORE);
        SetForegroundWindow(hwnd);
    }

    void commandWindowAction(const std::vector<std::string>& tokens, UINT message, int showMode) {
        if (tokens.size() < 2) {
            throw std::runtime_error(std::string(tokens[0]) + " 需要标题");
        }
        HWND hwnd = findWindowByTitle(utf8ToWide(join(tokens, 1)));
        if (hwnd == nullptr) {
            throw std::runtime_error("未找到窗口: " + join(tokens, 1));
        }
        if (message != 0) {
            PostMessageW(hwnd, message, 0, 0);
        } else {
            ShowWindow(hwnd, showMode);
        }
    }

    void commandDownload(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4 || tokens[2] != "到") {
            throw std::runtime_error("下载 语法: 下载 网址 到 路径");
        }
        const std::string command =
            "powershell -Command \""
            "$ProgressPreference='SilentlyContinue'; "
            "Invoke-WebRequest -UseBasicParsing -Uri '" + escapePowerShellSingleQuoted(tokens[1]) +
            "' -OutFile '" + escapePowerShellSingleQuoted(tokens[3]) + "'\"";
        if (std::system(command.c_str()) != 0) {
            throw std::runtime_error("下载失败: " + tokens[1]);
        }
    }

    void commandReadWeb(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4 || tokens[2] != "到") {
            throw std::runtime_error("读取网页 语法: 读取网页 网址 到 变量名");
        }
        setVariableValue(tokens[3], downloadUrlTextViaPowerShell(tokens[1]));
    }

    void commandRequest(const std::vector<std::string>& tokens) {
        if (tokens.size() < 5) {
            throw std::runtime_error("请求 语法: 请求 GET 网址 到 变量名 或 请求 POST 网址 数据 到 变量名");
        }
        if (tokens[1] == "GET") {
            if (tokens[3] != "到") {
                throw std::runtime_error("请求 GET 语法: 请求 GET 网址 到 变量名");
            }
            setVariableValue(tokens[4], requestUrlViaPowerShell("GET", tokens[2], ""));
            return;
        }
        if (tokens[1] == "POST") {
            if (tokens.size() < 6 || tokens[4] != "到") {
                throw std::runtime_error("请求 POST 语法: 请求 POST 网址 数据 到 变量名");
            }
            setVariableValue(tokens[5], requestUrlViaPowerShell("POST", tokens[2], tokens[3]));
            return;
        }
        throw std::runtime_error("当前仅支持 请求 GET 或 POST");
    }

    void commandDetectNetwork(const std::vector<std::string>& tokens) {
        if (tokens.size() < 3 || tokens[1] != "到") {
            throw std::runtime_error("检测网络 语法: 检测网络 到 变量名");
        }
        setVariableValue(tokens[2], detectNetworkViaPowerShell() ? "真" : "假");
    }

    void commandFindImage(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4 || tokens[2] != "到") {
            throw std::runtime_error("找图 语法: 找图 图片路径 到 变量名");
        }
        const ImageData image = loadImageFile(tokens[1]);
        setVariableValue(tokens[3], findImageOnScreen(image));
    }

    void commandOcr(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4 || tokens[2] != "到") {
            throw std::runtime_error("识字 语法: 识字 (x1,y1)->(x2,y2) 到 变量名");
        }
        const auto range = parsePointRange(tokens[1]);
        const int left = std::min(range.first.x, range.second.x);
        const int right = std::max(range.first.x, range.second.x);
        const int top = std::min(range.first.y, range.second.y);
        const int bottom = std::max(range.first.y, range.second.y);
        const int width = right - left + 1;
        const int height = bottom - top + 1;
        if (width <= 0 || height <= 0) {
            throw std::runtime_error("识字 区域无效");
        }

        const std::filesystem::path tempImage =
            std::filesystem::temp_directory_path() / ("operation_script_ocr_" + std::to_string(GetTickCount()) + ".bmp");
        HDC screenDc = GetDC(nullptr);
        HDC memDc = CreateCompatibleDC(screenDc);
        HBITMAP bitmap = CreateCompatibleBitmap(screenDc, width, height);
        HGDIOBJ oldObject = SelectObject(memDc, bitmap);
        BitBlt(memDc, 0, 0, width, height, screenDc, left, top, SRCCOPY);
        SelectObject(memDc, oldObject);

        try {
            saveBitmapToBmpFile(bitmap, memDc, tempImage);
        } catch (...) {
            DeleteObject(bitmap);
            DeleteDC(memDc);
            ReleaseDC(nullptr, screenDc);
            throw;
        }

        DeleteObject(bitmap);
        DeleteDC(memDc);
        ReleaseDC(nullptr, screenDc);

        std::vector<std::filesystem::path> candidates = {
            std::filesystem::path(currentScriptPath_).parent_path() / "tools" / "ocr.ps1",
            getExecutableDirectory() / "tools" / "ocr.ps1",
            std::filesystem::current_path() / "tools" / "ocr.ps1"
        };
        std::filesystem::path scriptPath;
        for (const auto& candidate : candidates) {
            if (std::filesystem::exists(candidate)) {
                scriptPath = candidate;
                break;
            }
        }
        if (scriptPath.empty()) {
            std::filesystem::remove(tempImage);
            throw std::runtime_error("缺少 OCR 脚本: tools/ocr.ps1");
        }

        try {
            setVariableValue(tokens[3], runOcrViaPowerShell(tempImage, scriptPath));
        } catch (...) {
            std::filesystem::remove(tempImage);
            throw;
        }
        std::filesystem::remove(tempImage);
    }

    void commandPickColor(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4 || tokens[2] != "到") {
            throw std::runtime_error("取色 语法: 取色 (x,y) 到 变量名");
        }
        const Point point = parsePoint(tokens[1]);
        HDC dc = GetDC(nullptr);
        COLORREF color = GetPixel(dc, point.x, point.y);
        ReleaseDC(nullptr, dc);
        setVariableValue(tokens[3], colorToHex(color));
    }

    void commandCompareColor(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4) {
            throw std::runtime_error("比色 语法: 比色 (x,y) 颜色 到 变量名");
        }
        const Point point = parsePoint(tokens[1]);
        const std::string expected = normalizeColorText(tokens[2]);
        if (tokens.size() >= 5 && tokens[3] == "到") {
            HDC dc = GetDC(nullptr);
            COLORREF color = GetPixel(dc, point.x, point.y);
            ReleaseDC(nullptr, dc);
            setVariableValue(tokens[4], normalizeColorText(colorToHex(color)) == expected ? "真" : "假");
            return;
        }
        throw std::runtime_error("比色 语法: 比色 (x,y) 颜色 到 变量名");
    }

    void commandFindColor(const std::vector<std::string>& tokens) {
        if (tokens.size() < 5 || tokens[3] != "到") {
            throw std::runtime_error("找色 语法: 找色 颜色 (x1,y1)->(x2,y2) 到 变量名");
        }
        const std::string expected = normalizeColorText(tokens[1]);
        const auto range = parsePointRange(tokens[2]);
        const int left = std::min(range.first.x, range.second.x);
        const int right = std::max(range.first.x, range.second.x);
        const int top = std::min(range.first.y, range.second.y);
        const int bottom = std::max(range.first.y, range.second.y);

        HDC dc = GetDC(nullptr);
        bool found = false;
        std::string result = "false";
        for (int y = top; y <= bottom && !found; ++y) {
            for (int x = left; x <= right; ++x) {
                COLORREF color = GetPixel(dc, x, y);
                if (normalizeColorText(colorToHex(color)) == expected) {
                    result = "(" + std::to_string(x) + "," + std::to_string(y) + ")";
                    found = true;
                    break;
                }
            }
        }
        ReleaseDC(nullptr, dc);
        setVariableValue(tokens[4], result);
    }

    void commandSaveScreenshot(const std::vector<std::string>& tokens) {
        if (tokens.size() < 2) {
            throw std::runtime_error("保存截图 需要路径");
        }
        const int screenX = GetSystemMetrics(SM_XVIRTUALSCREEN);
        const int screenY = GetSystemMetrics(SM_YVIRTUALSCREEN);
        const int screenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
        const int screenHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);

        HDC screenDc = GetDC(nullptr);
        HDC memDc = CreateCompatibleDC(screenDc);
        HBITMAP bitmap = CreateCompatibleBitmap(screenDc, screenWidth, screenHeight);
        HGDIOBJ oldObject = SelectObject(memDc, bitmap);
        BitBlt(memDc, 0, 0, screenWidth, screenHeight, screenDc, screenX, screenY, SRCCOPY);
        SelectObject(memDc, oldObject);

        try {
            saveBitmapToBmpFile(bitmap, memDc, tokens[1]);
        } catch (...) {
            DeleteObject(bitmap);
            DeleteDC(memDc);
            ReleaseDC(nullptr, screenDc);
            throw;
        }

        DeleteObject(bitmap);
        DeleteDC(memDc);
        ReleaseDC(nullptr, screenDc);
    }

    void commandTime(const std::vector<std::string>& tokens, const char* format) {
        if (tokens.size() < 3 || tokens[1] != "到") {
            throw std::runtime_error(std::string(tokens[0]) + " 语法: " + tokens[0] + " 到 变量名");
        }
        setVariableValue(tokens[2], formatNow(format));
    }

    void commandFormatTime(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4 || tokens[2] != "到") {
            throw std::runtime_error("格式时间 语法: 格式时间 模板 到 变量名");
        }
        setVariableValue(tokens[3], formatNow(tokens[1].c_str()));
    }

    void commandCopyFile(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4 || tokens[2] != "到") {
            throw std::runtime_error("复制文件 语法: 复制文件 源 到 目标");
        }
        std::filesystem::copy_file(tokens[1], tokens[3], std::filesystem::copy_options::overwrite_existing);
    }

    void commandMoveFile(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4 || tokens[2] != "到") {
            throw std::runtime_error("移动文件 语法: 移动文件 源 到 目标");
        }
        std::filesystem::rename(tokens[1], tokens[3]);
    }

    void commandWriteFile(const std::vector<std::string>& tokens, bool append) {
        if (tokens.size() < 3) {
            throw std::runtime_error(std::string(tokens[0]) + " 需要路径和内容");
        }
        std::ofstream output(tokens[1], append ? (std::ios::binary | std::ios::app) : std::ios::binary);
        if (!output) {
            throw std::runtime_error("无法写入文件: " + tokens[1]);
        }
        output << join(tokens, 2);
    }

    void commandReadFile(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4 || tokens[2] != "到") {
            throw std::runtime_error("读取文件 语法: 读取文件 路径 到 变量名");
        }
        std::ifstream input(tokens[1], std::ios::binary);
        if (!input) {
            throw std::runtime_error("无法读取文件: " + tokens[1]);
        }
        std::ostringstream buffer;
        buffer << input.rdbuf();
        setVariableValue(tokens[3], buffer.str());
    }

    void commandListDir(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4 || tokens[2] != "到") {
            throw std::runtime_error("列出目录 语法: 列出目录 路径 到 变量名");
        }
        std::ostringstream out;
        bool first = true;
        for (const auto& entry : std::filesystem::directory_iterator(tokens[1])) {
            if (!first) {
                out << ";";
            }
            first = false;
            out << entry.path().filename().string();
        }
        setVariableValue(tokens[3], out.str());
    }

    void commandExists(const std::vector<std::string>& tokens, bool allowDirectory) {
        if (tokens.size() < 4 || tokens[2] != "到") {
            throw std::runtime_error(std::string(tokens[0]) + " 语法: " + tokens[0] + " 路径 到 变量名");
        }
        bool exists = std::filesystem::exists(tokens[1]);
        if (!allowDirectory) {
            exists = exists && std::filesystem::is_regular_file(tokens[1]);
        }
        setVariableValue(tokens[3], exists ? "真" : "假");
    }

    void commandFileSize(const std::vector<std::string>& tokens) {
        if (tokens.size() < 4 || tokens[2] != "到") {
            throw std::runtime_error("文件大小 语法: 文件大小 路径 到 变量名");
        }
        setVariableValue(tokens[3], std::to_string(std::filesystem::file_size(tokens[1])));
    }

    void commandHelp() const {
        std::cout << "OperationScript 命令目录" << std::endl;
        for (const auto& entry : catalog_) {
            std::cout << "[" << entry.category << "] "
                      << entry.name << " - "
                      << (entry.implemented ? "已实现" : "待实现")
                      << " | " << entry.syntax
                      << std::endl;
        }
    }
};

}  // namespace

int main(int argc, char* argv[]) {
    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);

    if (argc < 2) {
        std::cout << "用法: explain.exe <脚本.os>" << std::endl;
        std::cout << "示例: explain.exe demo.os" << std::endl;
        return 1;
    }

    try {
        Interpreter interpreter;
        return interpreter.runFile(argv[1]);
    } catch (const std::exception& ex) {
        std::cerr << ex.what() << std::endl;
        return 1;
    }
}
