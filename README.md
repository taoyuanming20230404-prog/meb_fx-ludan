# meb_fx-ludan

分销后台表单批量录单工具（Python + Selenium）。

## 项目说明

本项目用于从 Excel 批量读取客户数据，并自动在分销后台页面执行录单流程。

核心脚本：

- `fx_ludan.py`：主程序（浏览器自动化录单）
- `test_project_match.py`：项目关键词离线匹配测试

## 运行环境

- Windows 10/11
- Python 3.10+
- Google Chrome
- 与 Chrome 主版本匹配的 ChromeDriver

安装依赖：

```bash
pip install -r requirements.txt
```

## 快速开始

1. 安装依赖并准备 `chromedriver`。
2. 运行主程序：

```bash
python fx_ludan.py
```

3. 按提示登录后台并选择 Excel 文件执行批量录单。

## 输入文件要求

Excel（`.xlsx`）建议包含列：

- `号码`（必填）
- `项目`
- `城市`
- `微信`（可选）

## 输出文件

程序会在 Excel 同目录生成：

- `录单反馈+YYYYMMDD.xlsx`（同日重复运行自动生成 `_2`、`_3` 后缀）
- `失败列表+YYYYMMDD.txt`（同日重复运行自动生成 `_2`、`_3` 后缀）
- `日志YYYYMMDDHHMMSS.txt`

## 词库配置

- `std_keywords.txt`：标准项目关键词
- `synonyms.json`：同义词映射

## 打包

可使用：

```bash
build_windows.bat
```

## Git 更新

日常提交命令：

```bash
git add .
git commit -m "your message"
git push
```

