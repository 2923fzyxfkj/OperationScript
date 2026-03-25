#include <algorithm>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <sstream>
#include <stdexcept>
#include <string>
#include <cstdlib>

std::string escapeRcString(const std::string& text) {
    std::string result;
    for (char ch : text) {
        if (ch == '\\' || ch == '"') {
            result.push_back('\\');
        }
        result.push_back(ch);
    }
    return result;
}

namespace {

std::string readAll(const std::filesystem::path& path) {
    std::ifstream input(path, std::ios::binary);
    if (!input) {
        throw std::runtime_error("无法读取脚本: " + path.string());
    }
    std::ostringstream buffer;
    buffer << input.rdbuf();
    return buffer.str();
}

void writeAll(const std::filesystem::path& path, const std::string& content) {
    std::ofstream output(path, std::ios::binary);
    if (!output) {
        throw std::runtime_error("无法写入文件: " + path.string());
    }
    output << content;
}

std::string escapeRawString(const std::string& text) {
    std::string result = text;
    size_t pos = 0;
    while ((pos = result.find(")OSB", pos)) != std::string::npos) {
        result.replace(pos, 4, ")OS_B");
        pos += 5;
    }
    return result;
}

std::string buildLauncherSource(const std::string& scriptText) {
    std::ostringstream out;
    out << "#include <filesystem>\n";
    out << "#include <fstream>\n";
    out << "#include <iostream>\n";
    out << "#include <string>\n\n";
    out << "#define main operation_script_explain_entry\n";
    out << "#include \"explain.cpp\"\n";
    out << "#undef main\n\n";
    out << "static const char* kEmbeddedScript = R\"OSB(" << escapeRawString(scriptText) << ")OSB\";\n\n";
    out << "int main() {\n";
    out << "    try {\n";
    out << "        const std::filesystem::path temp = std::filesystem::temp_directory_path() / \"operation_script_embedded.os\";\n";
    out << "        std::ofstream output(temp, std::ios::binary);\n";
    out << "        output << kEmbeddedScript;\n";
    out << "        output.close();\n";
    out << "        char program[] = \"embedded.exe\";\n";
    out << "        std::string scriptPath = temp.string();\n";
    out << "        char* argv[] = {program, scriptPath.data(), nullptr};\n";
    out << "        return operation_script_explain_entry(2, argv);\n";
    out << "    } catch (const std::exception& ex) {\n";
    out << "        std::cerr << ex.what() << std::endl;\n";
    out << "        return 1;\n";
    out << "    }\n";
    out << "}\n";
    return out.str();
}

void generate(const std::filesystem::path& scriptPath, const std::filesystem::path& outputCpp) {
    writeAll(outputCpp, buildLauncherSource(readAll(scriptPath)));
}

std::string buildResourceSource(const std::filesystem::path& iconPath, const std::string& version) {
    std::ostringstream out;
    if (!iconPath.empty()) {
        out << "1 ICON \"" << escapeRcString(iconPath.string()) << "\"\n";
    }
    if (!version.empty()) {
        std::string versionComma = version;
        std::replace(versionComma.begin(), versionComma.end(), '.', ',');
        out << "1 VERSIONINFO\n";
        out << "FILEVERSION " << versionComma << "\n";
        out << "PRODUCTVERSION " << versionComma << "\n";
        out << "BEGIN\n";
        out << "  BLOCK \"StringFileInfo\"\n";
        out << "  BEGIN\n";
        out << "    BLOCK \"080404B0\"\n";
        out << "    BEGIN\n";
        out << "      VALUE \"FileVersion\", \"" << escapeRcString(version) << "\\0\"\n";
        out << "      VALUE \"ProductVersion\", \"" << escapeRcString(version) << "\\0\"\n";
        out << "      VALUE \"ProductName\", \"OperationScript\\0\"\n";
        out << "    END\n";
        out << "  END\n";
        out << "  BLOCK \"VarFileInfo\"\n";
        out << "  BEGIN\n";
        out << "    VALUE \"Translation\", 0x804, 1200\n";
        out << "  END\n";
        out << "END\n";
    }
    return out.str();
}

void compile(const std::filesystem::path& scriptPath, const std::filesystem::path& outputExe,
             const std::filesystem::path& iconPath, const std::string& version) {
    const std::filesystem::path absoluteOutputExe = std::filesystem::absolute(outputExe);
    const std::filesystem::path generatedCpp = outputExe.string() + ".cpp";
    generate(scriptPath, generatedCpp);

    std::string resourceObject;
    if (!iconPath.empty() || !version.empty()) {
        const std::filesystem::path rcPath = outputExe.string() + ".rc";
        resourceObject = outputExe.string() + ".res";
        writeAll(rcPath, buildResourceSource(iconPath, version));
        const std::string rcCommand =
            "windres \"" + rcPath.string() + "\" -O coff -o \"" + resourceObject + "\"";
        const int rcCode = std::system(rcCommand.c_str());
        std::filesystem::remove(rcPath);
        if (rcCode != 0) {
            std::filesystem::remove(generatedCpp);
            throw std::runtime_error("资源编译失败，windres 返回码: " + std::to_string(rcCode));
        }
    }

    std::string command =
        "g++ \"" + generatedCpp.string() +
        "\" -std=c++17 -I\"" + std::filesystem::current_path().string() +
        "\" -lshell32 -luser32 -lgdi32 -lgdiplus";
    if (!resourceObject.empty()) {
        command += " \"" + resourceObject + "\"";
    }
    command += " -o \"" + outputExe.string() + "\"";
    const int code = std::system(command.c_str());
    std::filesystem::remove(generatedCpp);
    if (!resourceObject.empty()) {
        std::filesystem::remove(resourceObject);
    }

    if (code != 0) {
        throw std::runtime_error("编译失败，g++ 返回码: " + std::to_string(code));
    }

    const std::filesystem::path runtimeToolSource = std::filesystem::current_path() / "tools" / "ocr.ps1";
    if (std::filesystem::exists(runtimeToolSource)) {
        const std::filesystem::path outputDirectory = absoluteOutputExe.parent_path();
        const std::filesystem::path runtimeToolDir = outputDirectory / "tools";
        const std::filesystem::path runtimeToolDestination = runtimeToolDir / "ocr.ps1";
        std::filesystem::create_directories(runtimeToolDir);
        if (std::filesystem::absolute(runtimeToolSource) != std::filesystem::absolute(runtimeToolDestination)) {
            std::filesystem::copy_file(
                runtimeToolSource,
                runtimeToolDestination,
                std::filesystem::copy_options::overwrite_existing);
        }
    }
}

}  // namespace

int main(int argc, char* argv[]) {
    if (argc < 4) {
        std::cout << "用法:\n";
        std::cout << "  os-build 生成 <脚本.os> <输出.cpp>\n";
        std::cout << "  os-build 编译 <脚本.os> <输出.exe> [图标.ico] [版本号]\n";
        return 1;
    }

    try {
        const std::string mode = argv[1];
        if (mode == "生成") {
            generate(argv[2], argv[3]);
            std::cout << "已生成源码: " << argv[3] << std::endl;
            return 0;
        }
        if (mode == "编译") {
            const std::filesystem::path iconPath = argc >= 5 ? std::filesystem::path(argv[4]) : std::filesystem::path();
            const std::string version = argc >= 6 ? argv[5] : "";
            compile(argv[2], argv[3], iconPath, version);
            std::cout << "已生成程序: " << argv[3] << std::endl;
            return 0;
        }
        throw std::runtime_error("模式必须是 生成 或 编译");
    } catch (const std::exception& ex) {
        std::cerr << ex.what() << std::endl;
        return 1;
    }
}
