using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using NPOI.SS.UserModel;

namespace AutoMaster
{
    class Excel
    {
        public class Info
        {
            public class VariableInfo
            {
                public string type;
                public string name;
                public List<string> values = new List<string>();

                public string GetTypeString()
                {
                    if (type == "string") return "std::string";
                    return type;
                }
                public void WriteBytes(Stream stream, int index)
                {
                    var s = values[index];
                    byte[] b = new byte[0];
                    if (type == "string") {
                        b = System.Text.Encoding.GetEncoding(932).GetBytes(s);
                        byte[] b2 = BitConverter.GetBytes(b.Length);
                        stream.Write(b2, 0, b2.Length);
                    } else if (type == "int") {
                        b = BitConverter.GetBytes(int.Parse(s));
                    } else if (type == "float") {
                        b = BitConverter.GetBytes(float.Parse(s));
                    } else if (type == "bool") {
                        b = BitConverter.GetBytes(bool.Parse(s));
                    }
                    stream.Write(b, 0, b.Length);
                }
            }

            string name = "";
            bool constEnable = false;
            List<VariableInfo> values = new List<VariableInfo>();

            public bool GetHeader(ISheet worksheet)
            {
                var row = worksheet.GetRow(0);
                if (row == null || worksheet.LastRowNum < 4) return false;

                Dictionary<string, string> data = row.Cells.Select(cell => {
                    try {
                        return cell.StringCellValue.Split(new char[] { '=' });
                    } catch {
                        return cell.NumericCellValue.ToString().Split(new char[] { '=' });
                    }
                }).Where(v => v.Length >= 2).ToDictionary(v => v[0], v => v[1]);
                if (!data.ContainsKey("Name"))
                    return false;

                name = data["Name"];
                constEnable = data.ContainsKey("Const") && bool.Parse(data["Const"]);

                return true;
            }

            public void GetBody(ISheet worksheet)
            {
                List<List<string>> temp = new List<List<string>>();
                int lastRow = worksheet.LastRowNum;
                for (int i = 1; i <= lastRow; i++) {
                    // コメント行は無視
                    if (i == 3) continue;

                    IRow row = worksheet.GetRow(i);
                    if (row == null) break;

                    var ss = row.Cells.Select(cell => { try { return cell.StringCellValue; } catch { return cell.NumericCellValue.ToString(); } }).ToList();
                    if (string.IsNullOrEmpty(ss[0])) break;

                    temp.Add(ss);
                }
                for (int x = 0; x < temp[0].Count; ++x) {
                    VariableInfo info = new VariableInfo() {
                        type = temp[0][x],
                        name = temp[1][x],
                    };
                    if (string.IsNullOrEmpty(info.type) || info.type == "none")
                        continue;

                    for (int y = 2; y < temp.Count; ++y) {
                        info.values.Add(temp[y][x]);
                    }
                    values.Add(info);
                }
            }

            /// <summary>
            /// 読み込みソースファイルの出力
            /// </summary>
            /// <param name="path">書き込み先のディレクトリ</param>
            public void OutputSource(string path)
            {
                Console.Write(string.Format("{0}.hpp 作成中・・・", name));
                try {
                    var variables = values.Select(v => string.Format("{0} {1};", v.GetTypeString(), v.name)).ToArray();
                    var readers = values.Select(v => string.Format("{0} = reader.Read{1}();", v.name, v.type.ToUpper()[0] + v.type.Substring(1))).ToArray();
                    string s =
@"#pragma once

#include <string>
#include ""utility/StreamReader.hpp""

namespace MasterData
{{
class {0}Data
{{
public :
    {1}

    void Load(StreamReader& reader)
    {{
        {2}
    }}
}};
}}
";
                    var ss = string.Format(s, name, string.Join("\r\n    ", variables), string.Join("\r\n        ", readers));
                    if (!File.Exists(path)) Directory.CreateDirectory(path);
                    File.WriteAllText(string.Format("{0}/{1}.hpp", path, name), ss, Encoding.GetEncoding(932));
                    Console.WriteLine("成功");
                } catch {
                    Console.WriteLine("失敗");
                }
            }

            /// <summary>
            /// バイナリの出力
            /// </summary>
            /// <param name="path"></param>
            public void OutputBinary(string path)
            {
                Console.Write(string.Format("{0}.dat 作成中・・・", name));
                try {
                    if (!File.Exists(path)) Directory.CreateDirectory(path);
                    FileStream stream = new FileStream(string.Format("{0}/{1}.dat", path, name), FileMode.Create, FileAccess.Write);
                    if (!constEnable) {
                        byte[] b = BitConverter.GetBytes(values[0].values.Count);
                        stream.Write(b, 0, b.Length);
                    }
                    for (int y = 0; y < values[0].values.Count; ++y) {
                        for (int x = 0; x < values.Count; ++x) {
                            values[x].WriteBytes(stream, y);
                        }
                    }
                    stream.Close();
                    Console.WriteLine("成功");
                } catch {
                    Console.WriteLine("失敗");
                }
            }

            /// <summary>
            /// Include出力
            /// </summary>
            /// <returns></returns>
            public string GetIncludeString()
            {
                return string.Format("#include \"master/{0}.hpp\"", name);
            }

            /// <summary>
            /// 変数出力
            /// </summary>
            /// <returns></returns>
            public string GetVariableString(bool hpp)
            {
                string hs = hpp ? "extern " : "";
                if (constEnable) {
                    return string.Format("    {0}{1}Data {1};\r\n", hs, name);
                } else if (values[0].type == "int") {
                    return string.Format("    {0}std::vector<{1}Data> {1};\r\n", hs, name);
                } else if (values[0].type == "string") {
                    return string.Format("    {0}std::map<std::string, {1}Data> {1};\r\n", hs, name);
                }
                return "";
            }

            /// <summary>
            /// Reloadの出力
            /// </summary>
            /// <returns></returns>
            public string GetReloadString()
            {
                if (constEnable) {
                    return string.Format("        LoadConst(path + \"/{0}.dat\", std::bind(&{0}Data::Load, &{0}, std::placeholders::_1));", name);
                } else if (values[0].type == "int") {
                    return string.Format("        LoadArray(path + \"/{0}.dat\", {0});", name);
                } else if (values[0].type == "string") {
                    return string.Format("        LoadMap(path + \"/{0}.dat\", {0});", name);
                }
                return "";
            }
        }

        /// <summary>
        /// Excelファイルの読み込み
        /// </summary>
        /// <param name="paths">パス</param>
        /// <returns>読み込み後のリスト</returns>
        static public List<Info> Load(string[] paths)
        {
            List<Info> ret = new List<Info>();

            foreach (var path in paths) {
                IWorkbook workbook = WorkbookFactory.Create(path);
                foreach (ISheet worksheet in workbook) {
                    // ヘッダー取得
                    Info info = new Info();
                    if (!info.GetHeader(worksheet))
                        continue;

                    info.GetBody(worksheet);
                    ret.Add(info);
                }
            }

            return ret;
        }
    }
}
