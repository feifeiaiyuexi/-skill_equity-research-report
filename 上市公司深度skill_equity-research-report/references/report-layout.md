# 四页 WPS 研报

如果用户有新版型或既有原稿，以该文件的实际字体和结构为准。首次制作时可以使用 [空白样式](../assets/report-shell.docx) 和 [DOCX 生成器](../scripts/build_report.py)。预设继承青岛啤酒和中际旭创两份已完成研报的版式风格。

| 页码 | 默认放置内容 |
| --- | --- |
| 1 | 公司标题、结论式副标题、必要身份信息。左栏放核心判断、盈利表及风险，右栏放研究观点、股票信息、行情小图和核心来源。页末写一条后续验证条件。 |
| 2 | 两栏历史财务与现金转化指标，下方用一张图和短文解释品牌、产业链、量价或利润到现金的变化。 |
| 3 | 盈利预测、依据、三种情景。空间允许时加入可逐项相加的经营利润或 FCFF 复算表。 |
| 4 | 估值拆分、敏感性、同业交叉检验或市场价格反推，列出能改变判断的证据。投研作品可在资料来源前记录 AI 复核实例。 |

四页是默认上限，用户另有要求时按其上限。需填补留白时优先加入新的量化判断或解释，尤其是首页右栏与情景表下方。页内拥挤时先去除重复文字或缩短来源标签，再调整行距和字号。

## 样式与图表

A4 纵向，左右边距各 34 pt，首页较高的页眉，章标题和表头用深红 `#BC131B`。中文为楷体_GB2312，拉丁字母与数字用 Tahoma。预设正文约 10.75 pt，表格与脚注约 8–9 pt，标题约 24 pt。首页约 65:35 双栏，栏间距约 12 pt，页脚用动态 PAGE 域。正文、财务和情景表保持可编辑。

图表采用白底、浅网格、少量颜色。深红为主序列，暖红橙及浅赭为对照色。同类比较用同色系，标清单位、具体期间、数据口径和来源。行情图需给复权口径，序列不足时可用已有数据制作业务图，或留出简洁的判断栏。用户提供的虚构训练数据可用于模拟件，在图题和页眉标明模拟性质。图优先展示趋势、拆分或区间，避免重复表格全部数字。

## 生成器接口

用包含 `metadata` 和 `pages` 的 JSON 清单生成 DOCX。先完成证据与模型，把已经核验的文案、显示值和图路径写入清单。渲染器只负责排版和可编辑输出。正文资料链接可直接点击。用户提供不同模板时，使用适合该原稿的编辑方法。

- `metadata.company` 为公司名与代码。可选 `author`、`as_of`、`series`、`report_kind`、`masthead`、`logo`。股票日期可用 `YYYY.MM.DD`。标识图仅在当前用途允许时提供。
- `pages` 为依序排列的页面，每页是一组块。文本块用 `{"type":"paragraph","lead":"判断。","text":"依据。"}`。另外支持 `title`、`subtitle`、`assignment`、`section`、`note`、`band`，这些块用 `text` 字段。
- `columns` 用 `blocks` 包含左右两个块列表，`fractions` 默认为 50:50，首页推荐 `[0.646,0.354]`。可选 `gap` 为磅值。
- `table` 用 `headers` 和 `rows`。可传 `fractions` 为各列相对宽度，`text_cols` 为正文中应左对齐的列序号（从 0 开始）。`kv` 用 `rows` 保存标签和值的二列组合。
- `image` 用 `path`，接受本地图表文件，路径相对清单或为绝对路径。`sources` 用 `items`，每项为标题和实际的 HTTP(S) 来源链接。

```sh
python "/skill-directory/scripts/build_report.py" "/workspace/report.json" "/workspace/公司深度研究.docx"
```

使用具备 `python-docx` 的工作区运行环境。默认生成新文件。图表可用当前环境中合适的绘图软件制作。清单的分页意图仍需用最终 DOCX 渲染检查。使用者更改内容后，重新确认列宽、页数与图题对应。外部模板要先核对旧文档残留的作者、旧公司、学校标识、修订记录和链接。

## 中文预览

无界面渲染器遇到中文方块或空白时，优先为本次预览配置字体。在支持 Fontconfig 的环境，可以把下面内容保存为临时 `fonts.conf`，将字体目录和可写缓存改为本机路径，再用 `FONTCONFIG_FILE` 指向它运行 DOCX 渲染工具。检查配置所选替代字体实际存在。此配置只影响预览，Word 文件仍声明模板所用字体。

```xml
<?xml version="1.0"?>
<fontconfig>
  <dir>/System/Library/Fonts</dir>
  <dir>/System/Library/Fonts/Supplemental</dir>
  <cachedir>/private/tmp/equity-report-font-cache</cachedir>
  <match target="pattern">
    <test qual="any" name="family"><string>楷体_GB2312</string></test>
    <edit name="family" mode="prepend" binding="strong"><string>Kaiti SC</string></edit>
  </match>
</fontconfig>
```

可加入机器上已有的 Word 或 WPS 字体目录。用最终文件检查中文、页数和裁切，报告实际预览环境。使用本地替代字形完成渲染时，原生 WPS 的换行和分页以打开成品时为准。
