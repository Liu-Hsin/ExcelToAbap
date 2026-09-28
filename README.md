# ZCL_EXCEL2ABAP

把 `.xlsx` 文件读成 ABAP 动态内表的工具类。

不依赖 SAP 标准的上传/导入功能，也不依赖 `ALSM_EXCEL_TO_INTERNAL_TABLE` 那套 OLE 方案——直接解开 OOXML 包，用 iXML 解析 `worksheet` / `sharedStrings` / `styles` 三个部件，最后拼出一张「列字母作字段名」的动态内表返回。

---

## 特性

- **纯服务端解析**：不需要前端、不需要 Excel 客户端、不需要 GUI，可在后台作业里跑。
- **动态结构**：字段名直接取 Excel 的列字母（`A`、`B`、…、`Z`、`AA`），调用方不用预先定义结构。
- **多工作表支持**：默认读第一个 sheet，也可以按名称指定。
- **数值格式化**：识别单元格格式（内置 + 自定义 `numFmt`），把日期/时间的序列号还原成 `YYYYMMDD` / `HHMMSS` / `YYYYMMDD HHMMSS`。
- **合并单元格展开**：把左上角的值填充到整个合并区域。
- **共享字符串 / 内联字符串 / 富文本**：三种字符串存储方式都能正确取值。
- **自带清理**：全空行会在返回前剔除。

---

## 依赖与前置条件

| 依赖 | 用途 |
| --- | --- |
| `CL_OPENXML_HELPER` | 读取本地文件为 `XSTRING` |
| `CL_XLSX_DOCUMENT` | 打开 xlsx 包，获取 workbook 各部件 |
| `CL_XLSX_WORKBOOKPART` | 访问 sheet 列表、sharedStrings、styles 部件 |
| `IF_IXML_*` / `CL_IXML` | SAP 标准 XML 解析框架 |

> `CL_XLSX_DOCUMENT` / `CL_OPENXML_HELPER` 属于内部/自研工具类，引入本类前请确认目标系统已存在。

---

## 快速开始

```abap
data lr_tab type ref to data.
field-symbols <lt_data> type standard table.

" 读第一个工作表
try.
    lr_tab = zcl_excel2abap=>upload_file( im_file_name = 'D:\temp\demo.xlsx' ).
  catch cx_root into data(lx_err).
    " 处理异常
endtry.

" 读指定工作表（sheet 名不区分大小写）
lr_tab = zcl_excel2abap=>upload_file( im_file_name  = 'D:\temp\demo.xlsx'
                                      im_sheet_name = 'Sheet2' ).

assign lr_tab->* to <lt_data>.
loop at <lt_data> assigning field-symbol(<ls_row>).
  assign component 'A' of structure <ls_row> to field-symbol(<lv_a>).
  write: / <lv_a>.
endloop.
```

### 接口

```abap
class-methods UPLOAD_FILE
  importing
    !IM_FILE_NAME  type STRING            " 文件路径（服务器侧可见的路径）
    !IM_SHEET_NAME type STRING optional   " sheet 名，不传则取第一个
  returning
    value(CT_TAB)  type ref to DATA       " 动态内表引用
  raising
    CX_ROOT.
```

类为 `FINAL`、构造方法私有，**只能通过 `UPLOAD_FILE` 入口调用**。

---

## 返回结构

返回一张标准内表，行类型是运行时拼出来的结构：

- **字段名** = 该列的列字母（大写），例如 `A`、`B`、`AA`。
- **字段类型** = 全部为 `STRING`，日期/时间已转成字符串，**调用方需自行做类型转换和校验**。
- **行数** = 数据中出现的最大行号，但全空行已被删除，**因此行号与 Excel 行号不再一一对应**。
- 若指定的 sheet 名不存在，返回**空表**（不报错）。

---

## 处理流程

```
UPLOAD_FILE
  └─ CONSTRUCTOR        加载 xlsx 包（CL_OPENXML_HELPER → CL_XLSX_DOCUMENT）
  └─ OPEND_XML
       ├─ PARSE_SHEET_NAMES     workbook.xml → sheet 名 + rId
       ├─ SHARED_STRINGS_PART   sharedStrings.xml → 字符串池（si 索引从 0 起）
       ├─ PARSE_STYLES          styles.xml → 每个 cellXfs 的格式类型（D / T / DT / 空）
       └─ WORKSHEET             按名称或序号定位 sheet → READ_SHEET
            └─ PARSE_NODE       递归遍历，MV_MODE = 'SHEET' / 'SHARED' 两种模式
  └─ CREATE_DYNAMIC_TAB   拼结构 → 建行 → 填值 → 展开合并 → 删空行
```

`PARSE_NODE` 是核心：以 `<c>`（单元格）和 `<si>`（字符串条目）为**聚合边界**，进入时清空 `MV_CUR_*` 缓存，递归填充，退出时落库（`APPEND_CELL` / `APPEND_SHARED_STRING`）。

`RESOLVE_CELL_VALUE` 负责取值收尾：

| 单元格类型 `t` | 处理 |
| --- | --- |
| `s` | 拿值当索引去共享字符串池查表 |
| `e` | 错误值（`#N/A`、`#REF!` 等）加前缀，得到 `【错误】#N/A` |
| `inlineStr` | 原样保留（`<is><t>` 已在递归中取到） |
| 空 / 其他 | 数字、布尔值原样保留 |

随后若单元格带样式，再按 `numFmt` 分类做日期时间换算：

- `D` → `YYYYMMDD`（序列号 + 1899-12-30）
- `T` → `HHMMSS`（小数部分 × 86400 秒）
- `DT` → `YYYYMMDD HHMMSS`
- 空 → 普通数字，不动

---

## 已知限制与注意事项

1. **字段全是 STRING**。金额、日期、数量都拿不到原始类型，调用方自己做 `CONV`。
2. **空行被删除**。`DELETE ... WHERE table_line IS INITIAL` 是整行全空才删；但数据中的空白分隔行会一起消失，行定位不能靠索引。
3. **公式不返回值**。代码中 `MV_CUR_FORMULA` 记录了 `<f>` 内容，但后续没有使用——读到的是公式的缓存计算值（`<v>`），不是公式本身。
5. **合并单元格是值复制，不是结构还原**。整个区域都会被填上左上角的值；左上角无值则整块跳过。
6. **只读单元格值**。字体、颜色、边框、行高列宽等样式信息除日期时间分类外一律忽略。
7. **性能**。全量 XML 树解析 + 动态结构，字段数 = 实际出现过的列数，大表（数万行 × 上百列）内存占用和耗时都会明显上升。
8. **文件名需带完整路径**，且是应用服务器可见的路径（`CL_OPENXML_HELPER=>LOAD_LOCAL_FILE` 语义）。
9. 异常统一通过 `CX_ROOT` 抛出，调用方需自行 `CATCH` 并区分处理。

---

## 主要方法一览

| 方法 | 可见性 | 说明 |
| --- | --- | --- |
| `UPLOAD_FILE` | public | 唯一对外入口 |
| `OPEND_XML` | private | 编排四个解析步骤 |
| `PARSE_SHEET_NAMES` | private | sheet 名 / rId 表 |
| `SHARED_STRINGS_PART` | private | 解析共享字符串池 |
| `PARSE_STYLES` | private | 解析内置与自定义 `numFmt`，输出格式类型表 |
| `WORKSHEET` | private | 定位目标 sheet 并读取 |
| `PARSE_NODE` | private | 递归解析核心 |
| `RESOLVE_CELL_VALUE` | private | 共享串查表、错误值标记、日期时间换算 |
| `CREATE_DYNAMIC_TAB` | private | 动态结构、建行填值、合并展开、删空行 |
| `PARSE_CELL_REF` / `COL_TO_LETTER` / `GET_COL_LETTERS` | private | `A1` ↔ 行列号互转 |
| `CLASSIFY_FMT` | private | 按 `formatCode` 判断 `D` / `T` / `DT` |
