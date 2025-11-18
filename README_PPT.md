# 美股牛市分析PPT生成器

本项目用于生成McKinsey风格的美股牛市分析演示文稿。

## 文件说明

```
.
├── create_excel_data.py           # Excel数据文件生成脚本
├── generate_ppt.py                # PPT生成主脚本
├── ppt_data.xlsx                  # 数据源Excel文件
└── 美股牛市分析_第23页.pptx      # 生成的PPT文件
```

## 功能特性

### 第23页：美股11轮牛市——制度红利+科技周期+货币宽松

**布局设计**：
- ✓ 时间轴设计（1933-2020五个关键节点）
- ✓ 三大驱动因素分析框（制度红利/科技创新/货币政策）
- ✓ 核心洞察黄色高亮框
- ✓ McKinsey配色方案
- ✓ 16:9标准尺寸，白底蓝字设计

**配色方案**：
- PRIMARY_BLUE (#002960): 主标题、表头、重要标注
- SECONDARY_BLUE (#0065BD): 图表主色、二级标题
- LIGHT_BLUE (#C9F0FF): 洞察框背景、辅助区域
- IKEA_YELLOW (#FFDB00): 强调、警示、推荐标记
- WHITE (#FFFFFF): 深蓝背景上的文字
- GRAY (#808080): 次要信息、注释

**时间轴节点**：
1. 1933 - 证券法体系（制度红利）
2. 1971 - 布雷顿森林体系瓦解（美元霸权）
3. 1995 - 互联网泡沫 + PC时代（科技创新）
4. 2009 - 金融危机后QE时代（货币扩张）
5. 2020 - 疫情后无限QE（科技+货币）

## 安装依赖

```bash
pip install python-pptx openpyxl pandas
```

## 使用方法

### 1. 生成Excel数据文件

```bash
python3 create_excel_data.py
```

输出：`ppt_data.xlsx`

### 2. 生成PPT

```bash
python3 generate_ppt.py
```

输出：`美股牛市分析_第23页.pptx`

## 自定义修改

### 修改数据内容

编辑 `ppt_data.xlsx` 中的以下工作表：
- **第23页_美股上涨驱动**：包含时间轴、内容框和洞察数据

### 修改样式

在 `generate_ppt.py` 中调整以下参数：

```python
# 配色方案
PRIMARY_BLUE = RGBColor(0, 41, 96)
SECONDARY_BLUE = RGBColor(0, 101, 189)
LIGHT_BLUE = RGBColor(201, 240, 255)
IKEA_YELLOW = RGBColor(255, 219, 0)

# 排版规范
TITLE_FONT_SIZE = Pt(18)
BODY_FONT_SIZE = Pt(11)
CAPTION_FONT_SIZE = Pt(9)
```

## 技术栈

- **python-pptx**: PPT生成库
- **openpyxl**: Excel读写库
- **pandas**: 数据处理库（可选）

## 版本历史

### v1.0.0 (2024-11-18)
- ✓ 实现第23页时间轴布局
- ✓ 添加McKinsey配色方案
- ✓ 支持Excel数据源
- ✓ 自动生成洞察框

## 待扩展功能

- [ ] 添加更多页面模板
- [ ] 支持图表生成
- [ ] 批量生成多页PPT
- [ ] 自定义主题配置文件

## 许可证

MIT License

---

**作者**: Claude
**创建日期**: 2024-11-18
**最后更新**: 2024-11-18
