using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;

namespace AutoMaster
{
    class Program
    {
        static void Main(string[] args)
        {
            string xlsDir = "";
            string binaryDir = "";
            string sourceDir = "";

            for (int i = 0; i < args.Length; ++i) {
                switch (args[i]) {
                    case "-x": xlsDir = args[++i];      break;
                    case "-b": binaryDir = args[++i];   break;
                    case "-s": sourceDir = args[++i];   break;
                    default :                           break;
                }
            }
            if (string.IsNullOrEmpty(xlsDir) || string.IsNullOrEmpty(binaryDir)) {
                Console.WriteLine("AutoMaster.exe [-x エクセルディレクトリ(必須)] [-b バイナリ出力ディレクトリ(必須)] [-s ソース出力ディレクトリ] ");
                return;
            }

            var ss = Directory.GetFiles(xlsDir, "*.xls*", SearchOption.AllDirectories);
            var infos = Excel.Load(ss);
            foreach (var info in infos) {
                info.OutputBinary(binaryDir);
            }

            if (!string.IsNullOrEmpty(sourceDir)) {
                foreach (var info in infos) {
                    info.OutputSource(sourceDir);
                }
                string path = Path.Combine(sourceDir, "MasterData");
                if (File.Exists(path)) File.Delete(path);

                Console.Write("MasterData.hpp 作成中・・・");
                try {
                    OutputMasterDataSource(infos, path);
                    Console.WriteLine("成功");
                } catch {
                    Console.WriteLine("失敗");
                }
            }
        }

        /// <summary>
        /// マスターデータのヘッダー出力
        /// </summary>
        /// <param name="infos">Excelを読み込んだ情報</param>
        /// <param name="path">出力パス</param>
        static void OutputMasterDataSource(List<Excel.Info> infos, string path)
        {
            string hs =
@"#pragma once

#include <string>
#include <vector>
#include <map>
{0}

namespace MasterData
{{
{1}

    void Reload(const std::string& path);
}}
";

            string cs =
@"#include ""MasterData.hpp""
#include ""utility/FileUtility.hpp""
#include ""utility/StreamReader.hpp""
#include <functional>

namespace MasterData
{{
{0}
    void LoadConst(const std::string& path, std::function<void(StreamReader&)> func) {{
        auto buffer = File::ReadAllBytes(path);
        auto reader = StreamReader(buffer);
        func(reader);
    }}

    template <class T>
    void LoadArray(const std::string& path, T& value) {{
        auto buffer = File::ReadAllBytes(path);
        auto reader = StreamReader(buffer);
        T().swap(value);
        value.resize(reader.ReadInt());
        for (size_t i = 0; i < value.size(); ++i) {{
            value[i].Load(reader);
        }}
    }}

    template <class T>
    void LoadMap(const std::string& path, T& value) {{
        auto buffer = File::ReadAllBytes(path);
        auto reader = StreamReader(buffer);
        auto length = reader.ReadInt();
        value.clear();
        for (int i = 0; i < length; ++i) {{
            auto key = reader.ReadStringNoSeek();
            value[key].Load(reader);
        }}
    }}

    void Reload(const std::string& path) {{
{1}
    }}
}}
";
            var includes = infos.Select(v => v.GetIncludeString()).ToArray();
            var variables = infos.Select(v => v.GetVariableString(true)).ToArray();
            var ss = string.Format(hs,
                string.Join("\r\n", includes),
                string.Join("", variables));
            File.WriteAllText(path + ".hpp", ss, Encoding.GetEncoding(932));

            variables = infos.Select(v => v.GetVariableString(false)).OrderBy(v => v).ToArray();
            var reloadeds = infos.Select(v => v.GetReloadString()).OrderBy(v => v).ToArray();
            ss = string.Format(cs,
                string.Join("", variables),
                string.Join("\r\n", reloadeds));
            File.WriteAllText(path + ".cpp", ss, Encoding.GetEncoding(932));
        }
    }
}
