# Data_statistics Python Script 使用说明书

`Data_statistics.py` 文件是一个用于从 `volume_results` 文件夹中读取数据并生成统计报告的 Python 脚本。该脚本会自动遍历 `volume_results` 文件夹中的所有子文件夹，读取每个子文件夹中的 `report.json` 和 `reports.json` 文件，提取其中的相关数据，并将这些数据录入到一个名为 `Data_statistics.xlsx` 的 Excel 文件中。最后一列会显示每个数据文件夹的路径超链接，方便用户快速访问。

## 1. 安装所需的 Python 包

### 1.1 安装依赖

该脚本依赖于以下 Python 库：

- `openpyxl`：用于处理 Excel 文件的创建和操作。

您可以通过以下命令安装所需的依赖包：

```bash
pip install openpyxl
````

### 1.2 安装 `requirements.txt` 文件中的依赖

如果您有 `requirements.txt` 文件，您可以使用以下命令一次性安装所有依赖：

```bash
pip install -r requirements.txt
```

**requirements.txt** 内容：

```txt
openpyxl
```

## 2. 脚本功能说明

### 2.1 输入文件结构

* **`Data_statistics.py`** 脚本应该放置在 `volume_results` 文件夹内。
* `volume_results` 文件夹内应包含多个子文件夹，每个子文件夹应该包含两个 JSON 文件：

  * **`report.json`**：该文件包含 `plate` 和 `timestamp` 数据。
  * **`reports.json`**：该文件包含 `length`、`width`、`height` 和 `value` 数据。

#### `report.json` 文件示例：

```json
{
    "measurement_id": "蓝粤S17MU2_1758100178",
    "plate": "蓝粤S17MU2",
    "timestamp": "Wed Sep 17 17:09:38 2025\n"
}
```

#### `reports.json` 文件示例：

```json
{
    "dimensions": {
        "length": 1.6326746637348584,
        "width": 1.4982866011089253,
        "height": 1.219393099843988,
        "unit": "meters"
    },
    "volume": {
        "value": 2.982897170619853,
        "unit": "meters",
        "confidence": 13.61773833718951
    }
}
```

### 2.2 脚本执行过程

* **遍历文件夹**：脚本会遍历 `volume_results` 文件夹下的所有子文件夹。
* **读取文件**：从每个子文件夹中读取 `report.json` 和 `reports.json` 文件，并提取相关数据。
* **生成 Excel 文件**：将提取到的数据（包括文件夹名称、`plate`、`timestamp`、`length`、`width`、`height`、`value`）写入一个名为 `Data_statistics.xlsx` 的 Excel 文件中。
* **添加超链接**：在 Excel 文件的最后一列，将显示一个超链接，指向该子文件夹的路径。

### 2.3 输出文件

脚本运行后，会生成一个名为 `Data_statistics.xlsx` 的 Excel 文件。该文件的结构如下：

| Folder Name | Plate    | Timestamp                | Length | Width | Height | Value | Folder Link             |
| ----------- | -------- | ------------------------ | ------ | ----- | ------ | ----- | ----------------------- |
| Folder1     | 蓝粤S17MU2 | Wed Sep 17 17:09:38 2025 | 1.63   | 1.50  | 1.22   | 2.98  | file:///path/to/Folder1 |
| Folder2     | 蓝粤S18MU3 | Wed Sep 18 12:15:22 2025 | 2.12   | 1.89  | 1.35   | 4.56  | file:///path/to/Folder2 |
| ...         | ...      | ...                      | ...    | ...   | ...    | ...   | ...                     |

* **Folder Name**：子文件夹名称
* **Plate**：从 `report.json` 文件提取的 `plate`
* **Timestamp**：从 `report.json` 文件提取的 `timestamp`
* **Length**、**Width**、**Height**、**Value**：从 `reports.json` 文件提取的测量数据
* **Folder Link**：指向文件夹路径的超链接，点击可以直接访问文件夹

### 2.4 格式化输出

* **居中对齐**：所有的单元格（包括表头和数据行）都进行了居中对齐。
* **边框**：每个单元格（包括表头和数据行）都添加了薄边框，确保数据的清晰显示。

## 3. 脚本使用方法

### 3.1 准备工作

1. **确保脚本位置正确**：将 `Data_statistics.py` 脚本放置在 `volume_results` 文件夹内。
2. **文件夹结构**：

   * 确保每个子文件夹中包含 `report.json` 和 `reports.json` 文件。

### 3.2 运行脚本

1. 打开命令行（Terminal 或 Command Prompt）。

2. 使用 `cd` 命令进入脚本所在的目录：

   ```bash
   cd /path/to/volume_results
   ```

3. 运行脚本：

   ```bash
   python Data_statistics.py
   ```

4. **查看输出**：运行脚本后，`Data_statistics.xlsx` 文件将被保存在脚本所在目录中。您可以使用 Excel 打开该文件查看结果。

## 4. 常见问题

### 4.1 如何调整列宽或格式？

您可以在脚本中调整列宽和其他格式。例如，修改 `ws.column_dimensions` 部分来更改列的宽度，或者通过 `openpyxl` 提供的其他样式功能进行格式化。

### 4.2 脚本运行时提示编码错误怎么办？

确保 `report.json` 和 `reports.json` 文件是 UTF-8 编码格式。如果遇到编码问题，可以尝试修改文件编码为 UTF-8。

### 4.3 生成的 Excel 文件在哪里？

生成的 `Data_statistics.xlsx` 文件将保存在您运行脚本的目录中。您可以直接在文件夹中查看并打开该文件。

### 4.4 如何运行在其他机器上？

1. 将 `Data_statistics.py` 脚本和相关文件夹一起拷贝到另一台机器上。
2. 确保目标机器已经安装了 Python 环境。
3. 使用 `pip install -r requirements.txt` 安装依赖包。
4. 运行脚本并查看生成的 Excel 文件。

## 5. 总结

`Data_statistics.py` 脚本能够自动提取 `volume_results` 文件夹下各个子文件夹中的数据，并生成一个详细的统计报告。通过此报告，您可以方便地查看每个文件夹的相关数据，包括尺寸、体积以及超链接，快速访问数据文件夹。


