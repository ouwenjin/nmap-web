# 🔒 Security Toolkit

合并工具集，提供 **Nmap XML → Excel 报告解析**、**Excel → Web 探测清单提取**。

支持交互菜单和命令行两种使用方式。

---

## ✨ 功能

* **Nmap XML → Excel**

  * 自动合并当前目录下的多个 `*.xml` 文件（Nmap 扫描结果）。
  * 解析并提取 IP / 端口 / 服务 / 状态。
  * 自动去重，标记常见危险端口。
  * 输出为格式化后的 Excel（默认 `端口调研表.xlsx`）。

* **Excel → Web 探测清单**

  * 从已有 Excel 报告中提取 `IP:端口` 列表。
  * 支持 IPv4 与 IPv6（IPv6 自动加 `[]`）。
  * 自动去重，输出为 `web探测.txt`。
  * 支持手动选择当前目录下的 Excel 文件。

* **完整流程**

  * 打印横幅 → 执行 Nmap XML 解析 → 打印横幅 → 提取 IP:端口。
  * 每个步骤都会打印横幅，清晰展示执行进度。

---

## 📂 文件结构

```
nmap-web.py   # 主程序
merge.log             # 日志输出（自动生成）
端口调研表.xlsx       # 解析 Nmap 结果后的 Excel 报告
web探测.txt           # 从 Excel 提取的 IP:端口清单
```

---

## 🚀 使用方法

### 1. 安装依赖

```bash
pip install pandas openpyxl tqdm
```

### 2. 运行交互菜单

```bash
python nmap-web.py
```

会出现菜单：

```
=== 安全工具集菜单 ===
1) 解析 Nmap XML (生成端口调研表.xlsx)
2) 提取 Excel IP:端口 (选择 Excel 文件)
3) 完整流程 (每个步骤都打印横幅)
```

* **选项 1**：合并并解析当前目录下的 `*.xml`，生成 Excel 报告。
* **选项 2**：选择一个 Excel 文件，提取 IP:端口，输出 `web探测.txt`。
* **选项 3**：完整流程（解析 + 提取），中间会打印横幅提示。

### 3. 使用命令行子命令

```bash

# 解析当前目录下的 XML → Excel
python nmap-web.py nmap -o my_ports.xlsx

# 从 Excel 提取 IP:端口 → TXT
python nmap-web.py extract -i my_ports.xlsx -o my_targets.txt
```

---

## ⚠️ 注意事项

* **Nmap 扫描结果**必须是 `-oX` 参数导出的 XML 格式。
* **Excel 格式**要求第一列为 IP，第二列为端口/协议。
* 本工具不会执行端口扫描，仅解析已有扫描结果。

---

## 🧑‍💻 作者信息

* 作者: **zhkali**
* 仓库:

  * [GitHub](https://github.com/ouwenjin/nmap-web)
  * [Gitee](https://gitee.com/zhkali/nmap-web)

---

## 📜 License

MIT License
