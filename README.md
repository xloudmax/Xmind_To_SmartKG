# KGbot_converter
# Markdown 转 Excel 脚本

## 概述
请现将xmind文件转换为markdown文件（xmind本体导出）
该脚本用于解析 Markdown 文件，提取文件结构并将其转换为 Excel 格式。它会识别标题、定义、命题等元素，并将它们组织成一个包含实体、关系等信息的 Excel 文件。

## 使用方法

1. 安装必要的 Python 库，如果尚未安装，可以运行以下命令：

    ```bash
    pip install pandas
    ```

2. 打开脚本文件 `import.py`，确保将以下内容更新为您的 Markdown 文件路径：

    ```python
    with open(r'C:\Users\23038\Desktop\Converter\形式系统.md', 'r', encoding='utf-8') as file:
    ```

3. 运行脚本：

    ```bash
    python import.py
    ```

4. 脚本将生成一个名为 `input.xlsx` 的 Excel 文件，包含提取的结构信息。

## 注意事项

- 请确保 Markdown 文件的格式正确，以保证脚本的正确运行。
- 脚本中的关系类型默认为 '继承'，您可以根据需要进行修改。

# Excel 转换脚本

## 概述
注意！因为xlrd只支持xls格式，所以需要在Excel软件中，通过另存为转换xlsx->xls

该脚本用于将 Excel 文件转换为新的格式，包含实体和关系信息。它支持读取包含实体信息、关系信息和双向关系信息的 Excel 文件，并生成一个新的 Excel 文件。

## 使用方法

1. 安装必要的 Python 库，如果尚未安装，可以运行以下命令：

    ```bash
    pip install xlrd xlsxwriter
    ```

2. 打开脚本文件 `excel_converter.py`，确保将以下内容更新为您的 Excel 文件路径：

    ```python
    inputfile = r'C:\Users\23038\Desktop\Converter\input.xls'
    outputfile = r'C:\Users\23038\Desktop\Converter\output.xlsx'
    ```

3. 运行脚本：

    ```bash
    python excel_converter.py
    ```

4. 脚本将读取输入文件，转换为新的格式，并生成输出文件。

## 注意事项

- 脚本默认读取包含实体信息的第一个 sheet、关系信息的第二个 sheet，以及包含双向关系信息的第三个 sheet。
- 输出文件将包含两个 sheet，一个用于实体信息，另一个用于关系信息。
