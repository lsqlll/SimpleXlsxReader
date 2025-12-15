# C++表格读取程序框架设计文档

## 1. 概述

本文档定义了一个基于**策略模式**与**工厂模式**的C++表格读取库框架，旨在提供类似Python `openpyxl`的灵活、易用的接口。框架支持多种表格格式（XLS、XLSX、CSV等），遵循SOLID原则，具有良好的可扩展性。

### 1.1 设计目标

- **统一接口**：通过抽象接口访问不同格式的表格文件
- **格式解耦**：通过策略模式隔离不同格式的解析逻辑
- **易于扩展**：通过工厂模式简化新格式的集成
- **资源安全**：通过RAII和智能指针管理C库资源
- **类型安全**：通过强类型接口避免运行时错误

### 1.2 核心设计模式

1. **策略模式（Strategy Pattern）**：定义不同格式的读取策略
2. **工厂模式（Factory Pattern）**：根据文件类型创建相应的读取器
3. **RAII模式**：自动管理底层C资源的生命周期

---

## 2. 核心类与接口定义

### 2.1 单元格接口层次

#### 2.1.1 ICell - 单元格抽象接口

```cpp
class ICell {
public:
    virtual ~ICell() = default;

    virtual CellType getType() const = 0;

    virtual std::size_t getRow() const = 0;
    virtual std::size_t getCol() const = 0;
    virtual CellPosition getPosition() const = 0;

    virtual std::string asString(bool trimWhitespace = false) const = 0;
    virtual double asDouble() const = 0;
    virtual bool asBool() const = 0;

    virtual bool isBlank() const = 0;
    virtual bool isNumber() const = 0;
    virtual bool isString() const = 0;
    virtual bool isBool() const = 0;
    virtual bool isDate() const = 0;
};
```

**职责**：

- 定义单元格的通用接口
- 提供坐标、类型、值访问的统一抽象
- 支持类型查询与转换

#### 2.1.2 XlsCell - XLS格式单元格实现

现有`XlsCell`类需要实现`ICell`接口，成为具体实现：

```cpp
class XlsCell : public ICell {
private:
    std::shared_ptr<xls::xlsCell> cell_;
    CellPosition location_;
    std::optional<CellType> type_;
    std::variant<std::monostate, std::string, double, bool> value_;

    void inferValue(bool trimWs);

public:
    explicit XlsCell(xls::xlsCell* cell);

    CellType getType() const override;
    std::size_t getRow() const override;
    std::size_t getCol() const override;
    CellPosition getPosition() const override;

    std::string asString(bool trimWhitespace = false) const override;
    double asDouble() const override;
    bool asBool() const override;

    bool isBlank() const override;
    bool isNumber() const override;
    bool isString() const override;
    bool isBool() const override;
    bool isDate() const override;
};
```

**适配说明**：

- 保留现有的类型推断逻辑
- 添加`ICell`接口的实现方法
- 现有的`asStdString()`改为`asString()`
- 现有的`asLogical()`改为`asBool()`

---

### 2.2 工作表接口层次

#### 2.2.1 IWorksheet - 工作表抽象接口

```cpp
class IWorksheet {
public:
    virtual ~IWorksheet() = default;

    virtual std::string getName() const = 0;
    virtual std::size_t getIndex() const = 0;

    virtual std::pair<std::size_t, std::size_t> getDimensions() const = 0;
    virtual std::size_t getMaxRow() const = 0;
    virtual std::size_t getMaxColumn() const = 0;

    virtual std::shared_ptr<ICell> getCell(std::size_t row, std::size_t col) = 0;
    virtual std::shared_ptr<ICell> getCell(const std::string& address) = 0;
    virtual std::shared_ptr<ICell> getCell(const CellPosition& position) = 0;

    virtual std::vector<std::shared_ptr<ICell>> getRow(std::size_t rowIndex) = 0;
    virtual std::vector<std::shared_ptr<ICell>> getColumn(std::size_t colIndex) = 0;

    virtual bool isLoaded() const = 0;
    virtual void ensureLoaded() = 0;
};
```

**职责**：

- 提供工作表的元数据访问（名称、索引、维度）
- 支持多种方式访问单元格（行列索引、Excel地址、`CellPosition`）
- 支持批量读取（按行、按列）
- 延迟加载支持（`isLoaded()`/`ensureLoaded()`）

#### 2.2.2 XlsWorksheet - XLS格式工作表实现

```cpp
class XlsWorksheet : public IWorksheet {
private:
    std::size_t index_;
    std::string name_;
    xls::xlsWorkSheet* sheet_;
    bool isLoaded_;
    std::size_t maxRow_;
    std::size_t maxCol_;

public:
    XlsWorksheet(std::size_t index, std::string name, xls::xlsWorkSheet* sheet);
    ~XlsWorksheet() override;

    std::string getName() const override;
    std::size_t getIndex() const override;

    std::pair<std::size_t, std::size_t> getDimensions() const override;
    std::size_t getMaxRow() const override;
    std::size_t getMaxColumn() const override;

    std::shared_ptr<ICell> getCell(std::size_t row, std::size_t col) override;
    std::shared_ptr<ICell> getCell(const std::string& address) override;
    std::shared_ptr<ICell> getCell(const CellPosition& position) override;

    std::vector<std::shared_ptr<ICell>> getRow(std::size_t rowIndex) override;
    std::vector<std::shared_ptr<ICell>> getColumn(std::size_t colIndex) override;

    bool isLoaded() const override;
    void ensureLoaded() override;
};
```

**设计要点**：

- 持有底层C库的`xlsWorkSheet*`指针
- 延迟解析工作表内容（首次访问时调用`xls_parseWorkSheet`）
- 析构时负责释放工作表资源（如果需要）

---

### 2.3 工作簿接口层次

#### 2.3.1 IWorkbook - 工作簿抽象接口

```cpp
class IWorkbook {
public:
    virtual ~IWorkbook() = default;

    virtual bool isOpen() const = 0;

    virtual std::size_t getSheetCount() const = 0;
    virtual std::vector<std::string> getSheetNames() const = 0;

    virtual std::shared_ptr<IWorksheet> getSheet(std::size_t index) = 0;
    virtual std::shared_ptr<IWorksheet> getSheet(const std::string& name) = 0;

    virtual std::shared_ptr<IWorksheet> getActiveSheet() = 0;

    virtual std::filesystem::path getFilePath() const = 0;
};
```

**职责**：

- 代表整个表格文件
- 管理工作表集合
- 提供按索引或名称访问工作表的方法
- 跟踪文件打开状态

#### 2.3.2 XlsWorkbook - XLS格式工作簿实现

```cpp
class XlsWorkbook : public IWorkbook {
private:
    std::filesystem::path filePath_;
    std::unique_ptr<xls::xlsWorkBook, decltype(&xls::xls_close_WB)> workbook_;
    std::vector<std::shared_ptr<XlsWorksheet>> sheets_;
    bool isOpen_;

    void loadSheetMetadata();

public:
    explicit XlsWorkbook(const std::filesystem::path& path);
    ~XlsWorkbook() override = default;

    bool isOpen() const override;

    std::size_t getSheetCount() const override;
    std::vector<std::string> getSheetNames() const override;

    std::shared_ptr<IWorksheet> getSheet(std::size_t index) override;
    std::shared_ptr<IWorksheet> getSheet(const std::string& name) override;

    std::shared_ptr<IWorksheet> getActiveSheet() override;

    std::filesystem::path getFilePath() const override;
};
```

**设计要点**：

- 使用自定义删除器的`unique_ptr`管理`xlsWorkBook*`
- 构造时打开文件并加载工作表元数据
- 缓存`XlsWorksheet`对象避免重复创建
- 遵循RAII，析构时自动关闭工作簿

---

### 2.4 策略模式接口

#### 2.4.1 IReaderStrategy - 读取策略抽象接口

```cpp
class IReaderStrategy {
public:
    virtual ~IReaderStrategy() = default;

    virtual std::unique_ptr<IWorkbook> openWorkbook(const std::filesystem::path& path) = 0;

    virtual bool canHandle(const std::filesystem::path& path) const = 0;

    virtual std::string getFormatName() const = 0;
};
```

**职责**：

- 定义读取表格文件的通用接口
- 封装特定格式的打开逻辑
- 提供格式兼容性检查

#### 2.4.2 XlsReaderStrategy - XLS格式读取策略

```cpp
class XlsReaderStrategy : public IReaderStrategy {
public:
    std::unique_ptr<IWorkbook> openWorkbook(const std::filesystem::path& path) override;

    bool canHandle(const std::filesystem::path& path) const override;

    std::string getFormatName() const override;
};
```

**实现说明**：

- `openWorkbook()`：创建并返回`XlsWorkbook`实例
- `canHandle()`：检查文件扩展名是否为`.xls`
- `getFormatName()`：返回`"XLS"`

#### 2.4.3 XlsxReaderStrategy - XLSX格式读取策略（桩）

```cpp
class XlsxReaderStrategy : public IReaderStrategy {
public:
    std::unique_ptr<IWorkbook> openWorkbook(const std::filesystem::path& path) override;

    bool canHandle(const std::filesystem::path& path) const override;

    std::string getFormatName() const override;
};
```

**实现说明**：

- 当前为桩实现，`openWorkbook()`抛出未实现异常
- `canHandle()`检查`.xlsx`扩展名
- 后续集成xlnt库时完成实现

#### 2.4.4 CsvReaderStrategy - CSV格式读取策略（桩）

```cpp
class CsvReaderStrategy : public IReaderStrategy {
public:
    std::unique_ptr<IWorkbook> openWorkbook(const std::filesystem::path& path) override;

    bool canHandle(const std::filesystem::path& path) const override;

    std::string getFormatName() const override;
};
```

**实现说明**：

- 当前为桩实现
- `canHandle()`检查`.csv`扩展名
- 后续根据需求完成CSV解析实现

---

### 2.5 工厂模式类

#### 2.5.1 ReaderStrategyFactory - 策略工厂

```cpp
class ReaderStrategyFactory {
private:
    std::vector<std::unique_ptr<IReaderStrategy>> strategies_;

    ReaderStrategyFactory();

public:
    static ReaderStrategyFactory& getInstance();

    std::unique_ptr<IReaderStrategy> createStrategy(const std::filesystem::path& path);

    void registerStrategy(std::unique_ptr<IReaderStrategy> strategy);

    std::vector<std::string> getSupportedFormats() const;
};
```

**职责**：

- 管理所有注册的读取策略
- 根据文件路径选择合适的策略
- 单例模式保证全局唯一

**工作流程**：

1. 构造时注册所有内置策略（XLS、XLSX、CSV）
2. `createStrategy()`遍历策略列表，调用`canHandle()`找到匹配策略
3. 返回匹配策略的副本（通过克隆或工厂方法）

#### 2.5.2 WorkbookFactory - 工作簿工厂

```cpp
class WorkbookFactory {
public:
    static std::unique_ptr<IWorkbook> open(const std::filesystem::path& path);

    static std::unique_ptr<IWorkbook> openWithStrategy(
        const std::filesystem::path& path,
        std::unique_ptr<IReaderStrategy> strategy
    );
};
```

**职责**：

- 提供统一的工作簿创建接口
- 整合策略选择与工作簿实例化

**使用示例**：

```cpp
auto workbook = WorkbookFactory::open("data.xls");
auto sheet = workbook->getSheet(0);
auto cell = sheet->getCell("A1");
std::cout << cell->asString() << std::endl;
```

---

## 3. 类关系图（文本描述）

```
┌─────────────────────────────────────────────────────────────┐
│                      用户代码层                              │
│  WorkbookFactory::open("file.xls") → IWorkbook              │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│                      工厂层                                  │
│  WorkbookFactory ──uses──> ReaderStrategyFactory            │
│                              │                               │
│                              ▼                               │
│                        IReaderStrategy                       │
│                       ┌───────┴───────┐                      │
│              XlsReaderStrategy  XlsxReaderStrategy           │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│                      工作簿层                                │
│                      IWorkbook                               │
│                  ┌────────┴────────┐                         │
│            XlsWorkbook       XlsxWorkbook                    │
│                  │                                            │
│                  │ owns                                       │
│                  ▼                                            │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│                      工作表层                                │
│                     IWorksheet                               │
│                  ┌────────┴────────┐                         │
│            XlsWorksheet      XlsxWorksheet                   │
│                  │                                            │
│                  │ creates                                    │
│                  ▼                                            │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│                      单元格层                                │
│                       ICell                                  │
│                  ┌────────┴────────┐                         │
│                XlsCell       XlsxCell                        │
└─────────────────────────────────────────────────────────────┘
```

---

## 4. 交互流程

### 4.1 打开工作簿的完整流程

```
用户代码
  │
  ├─→ WorkbookFactory::open("data.xls")
      │
      ├─→ ReaderStrategyFactory::getInstance().createStrategy("data.xls")
      │    │
      │    ├─→ XlsReaderStrategy::canHandle("data.xls") → true
      │    └─→ 返回 XlsReaderStrategy 实例
      │
      ├─→ XlsReaderStrategy::openWorkbook("data.xls")
      │    │
      │    ├─→ 调用 xls_open() 打开文件
      │    ├─→ 创建 XlsWorkbook 实例
      │    └─→ XlsWorkbook::loadSheetMetadata()
      │         │
      │         └─→ 为每个工作表创建 XlsWorksheet 对象（未解析内容）
      │
      └─→ 返回 unique_ptr<IWorkbook>
```

### 4.2 访问单元格的流程

```
用户代码
  │
  ├─→ workbook->getSheet(0)
  │    └─→ 返回 XlsWorksheet 实例（shared_ptr<IWorksheet>）
  │
  ├─→ sheet->getCell("A1")
      │
      ├─→ XlsWorksheet::ensureLoaded()
      │    └─→ 如果 !isLoaded_，调用 xls_parseWorkSheet()
      │
      ├─→ 解析地址 "A1" → (row=0, col=0)
      │
      ├─→ 调用 xls_cell(sheet_, 0, 0)
      │
      ├─→ 创建 XlsCell 实例
      │
      └─→ 返回 shared_ptr<ICell>
```

---

## 5. 关键设计决策

### 5.1 依赖倒置原则

- **高层模块**（`WorkbookFactory`）依赖抽象接口（`IReaderStrategy`、`IWorkbook`）
- **低层模块**（`XlsReaderStrategy`、`XlsWorkbook`）实现抽象接口
- **好处**：高层逻辑无需关心具体格式，易于测试和维护

### 5.2 开闭原则

- **对扩展开放**：添加新格式只需：
  1. 实现`IReaderStrategy`（如`OdsReaderStrategy`）
  2. 实现`IWorkbook`、`IWorksheet`、`ICell`
  3. 在`ReaderStrategyFactory`中注册新策略
- **对修改封闭**：无需修改现有代码（工厂、接口、用户代码）

### 5.3 资源管理

- **RAII**：`XlsWorkbook`的析构函数自动调用`xls_close_WB()`
- **智能指针**：
  - `unique_ptr`用于独占所有权（工作簿、策略）
  - `shared_ptr`用于共享所有权（工作表、单元格）
- **自定义删除器**：`unique_ptr<xls::xlsWorkBook, decltype(&xls::xls_close_WB)>`

### 5.4 延迟加载

- **工作表内容延迟解析**：只有在首次访问单元格时才调用`xls_parseWorkSheet()`
- **好处**：减少内存占用，提升打开速度

### 5.5 类型安全

- **强类型接口**：返回`shared_ptr<ICell>`而非原始指针
- **异常安全**：构造失败时抛出明确的异常（`FileNotFoundException`等）

---

## 6. 扩展点

### 6.1 添加新表格格式

**步骤**：

1. **实现策略接口**：
   
   ```cpp
   class OdsReaderStrategy : public IReaderStrategy {
       std::unique_ptr<IWorkbook> openWorkbook(const std::filesystem::path& path) override;
       bool canHandle(const std::filesystem::path& path) const override;
       std::string getFormatName() const override;
   };
   ```

2. **实现工作簿、工作表、单元格**：
   
   ```cpp
   class OdsWorkbook : public IWorkbook { /* ... */ };
   class OdsWorksheet : public IWorksheet { /* ... */ };
   class OdsCell : public ICell { /* ... */ };
   ```

3. **注册策略**：
   
   ```cpp
   ReaderStrategyFactory::getInstance().registerStrategy(
       std::make_unique<OdsReaderStrategy>()
   );
   ```

**无需修改的模块**：

- `WorkbookFactory`
- `IWorkbook`/`IWorksheet`/`ICell`接口
- 用户代码

### 6.2 添加新单元格类型

如果需要支持新的数据类型（如`DATETIME`、`FORMULA`）：

1. 扩展`CellType`枚举：
   
   ```cpp
   enum class CellType : uint8_t {
       STRING = 0,
       NUMBER,
       BOOL,
       UNKNOWN,
       BLANK,
       DATE,
       DATETIME,  // 新增
       FORMULA    // 新增
   };
   ```

2. 在`ICell`中添加新方法：
   
   ```cpp
   virtual std::chrono::system_clock::time_point asDateTime() const = 0;
   virtual std::string getFormula() const = 0;
   ```

3. 各具体单元格类实现新方法

---

## 7. 与现有代码的整合

### 7.1 需要修改的文件

| 文件           | 变更类型 | 主要变更内容                                                                    |
| ------------ | ---- | ------------------------------------------------------------------------- |
| `XlsCell.h`  | 修改   | 继承`ICell`，实现接口方法，重命名部分方法                                                  |
| `strategy.h` | 重构   | 删除现有`ReadStrategy`，重新实现`IReaderStrategy`及其子类                              |
| `reader.h`   | 重构   | 删除现有`TableReader`，实现`IWorkbook`/`IWorksheet`/`XlsWorkbook`/`XlsWorksheet` |
| `factory.h`  | 新增   | 实现`ReaderStrategyFactory`和`WorkbookFactory`                               |
| `type.h`     | 修改   | 删除`XLSheets`/`XLSRow`，可能需要调整类型别名                                          |

### 7.2 需要新增的文件

| 文件                                              | 内容                    |
| ----------------------------------------------- | --------------------- |
| `ICell.h`                                       | `ICell`接口定义           |
| `IWorksheet.h`                                  | `IWorksheet`接口定义      |
| `IWorkbook.h`                                   | `IWorkbook`接口定义       |
| `IReaderStrategy.h`                             | `IReaderStrategy`接口定义 |
| `XlsWorkbook.h` / `XlsWorkbook.cpp`             | `XlsWorkbook`实现       |
| `XlsWorksheet.h` / `XlsWorksheet.cpp`           | `XlsWorksheet`实现      |
| `XlsReaderStrategy.h` / `XlsReaderStrategy.cpp` | `XlsReaderStrategy`实现 |
| `WorkbookFactory.h` / `WorkbookFactory.cpp`     | 工厂类实现                 |

### 7.3 保留的现有组件

- `CellType.h`：`CellType`枚举和`CellPosition`结构体
- `utils.h`：工具函数（trim、isEmpty、parseAddress等）
- `Exceptions.h`：异常类定义
- `ResourceManager.h`：RAII资源管理模板（修复编译错误后）

---

## 8. 接口易用性验证

### 8.1 用例1：读取单个单元格

```cpp
auto workbook = WorkbookFactory::open("sales.xls");
auto sheet = workbook->getSheet("Q1 Sales");
auto cell = sheet->getCell("B5");
std::cout << "Value: " << cell->asString() << std::endl;
```

**对比openpyxl**：

```python
workbook = openpyxl.load_workbook("sales.xlsx")
sheet = workbook["Q1 Sales"]
cell = sheet["B5"]
print("Value:", cell.value)
```

### 8.2 用例2：遍历工作表

```cpp
auto workbook = WorkbookFactory::open("data.xls");
for (const auto& sheetName : workbook->getSheetNames()) {
    auto sheet = workbook->getSheet(sheetName);
    std::cout << "Sheet: " << sheetName 
              << " (" << sheet->getMaxRow() << " rows)" << std::endl;
}
```

**对比openpyxl**：

```python
workbook = openpyxl.load_workbook("data.xlsx")
for sheetName in workbook.sheetnames:
    sheet = workbook[sheetName]
    print(f"Sheet: {sheetName} ({sheet.max_row} rows)")
```

### 8.3 用例3：读取整行

```cpp
auto workbook = WorkbookFactory::open("data.xls");
auto sheet = workbook->getSheet(0);
auto row = sheet->getRow(2);  // 第3行（0-indexed）

for (const auto& cell : row) {
    std::cout << cell->asString() << "\t";
}
std::cout << std::endl;
```

**对比openpyxl**：

```python
workbook = openpyxl.load_workbook("data.xlsx")
sheet = workbook.worksheets[0]
row = sheet[3]  # 第3行（1-indexed）

for cell in row:
    print(cell.value, end="\t")
print()
```

---

## 9. 性能与优化考虑

### 9.1 延迟加载

- 工作表内容仅在首次访问时解析
- 避免一次性加载所有工作表

### 9.2 智能指针开销

- `shared_ptr`有引用计数开销，但在多线程环境下更安全
- 对于性能敏感场景，可考虑对象池或内存池

### 9.3 字符串转换

- `XlsCell`中的字符串操作（trim、toLower）可能频繁调用
- 考虑缓存转换结果或使用`string_view`

---

## 10. 质量保证

### 10.1 编译时检查

- 所有接口方法使用`override`关键字
- 禁止拷贝（`= delete`）或显式允许（`= default`）
- 使用`[[nodiscard]]`标记返回值不应忽略的方法

### 10.2 运行时检查

- 异常规格明确（`FileNotFoundException`、`IndexOutOfRangeException`等）
- 边界检查（工作表索引、单元格坐标）

### 10.3 测试建议

- **单元测试**：每个接口实现类的独立测试
- **集成测试**：完整的打开→读取→转换流程
- **边界测试**：空文件、单单元格、大文件

---

## 11. 后续实施路线图

### 阶段1：接口定义（当前文档）

- ✅ 定义所有抽象接口
- ✅ 设计类关系图
- ✅ 验证接口易用性

### 阶段2：XLS格式完整实现

1. 实现`ICell`/`IWorksheet`/`IWorkbook`接口
2. 重构`XlsCell`适配`ICell`
3. 实现`XlsWorksheet`和`XlsWorkbook`
4. 实现`XlsReaderStrategy`
5. 实现`WorkbookFactory`
6. 修复`ResourceManager`编译错误
7. 编写单元测试

### 阶段3：XLSX格式实现

1. 集成xlnt库
2. 实现`XlsxCell`/`XlsxWorksheet`/`XlsxWorkbook`
3. 实现`XlsxReaderStrategy`
4. 编写测试

### 阶段4：CSV格式实现

1. 选择CSV解析库（或自实现）
2. 实现CSV相关类
3. 实现`CsvReaderStrategy`

### 阶段5：优化与扩展

1. 性能分析与优化
2. 添加写入功能（可选）
3. 添加公式计算支持（可选）

---

## 12. 总结

本框架通过**策略模式**与**工厂模式**的结合，实现了：

1. **统一的接口**：`IWorkbook`/`IWorksheet`/`ICell`提供类似openpyxl的使用体验
2. **格式解耦**：每种格式独立实现，互不影响
3. **易于扩展**：添加新格式只需实现接口并注册策略，无需修改现有代码
4. **资源安全**：通过RAII和智能指针自动管理C资源
5. **类型安全**：强类型接口减少运行时错误

该框架为后续开发提供了清晰的蓝图，满足了设计文档的所有验收标准。
