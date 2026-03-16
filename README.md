# 达成度报告生成器

> A Python tool to generate course achievement analysis reports from student grade Excel files.

从学生成绩单Excel文件自动生成达成度分析报告的工具。

**当前版本**: v1.7

## 功能特点

- 🖥️ **图形界面** - 基于 CustomTkinter 的现代化桌面应用
- 📊 **批量处理** - 支持一次处理多个成绩单文件
- 🔍 **智能识别** - 自动识别成绩单结构，支持多列并排数据
- 📈 **图表生成** - 自动生成达成度折线图和统计柱状图
- ⚠️ **完善的错误提示** - 实时验证配置，详细的错误信息
- 📁 **文件覆盖确认** - 同名文件支持覆盖、重命名或跳过

## 快速开始

### 方式一：使用桌面应用（推荐）

#### 运行应用
```bash
cd achievement_report_app
python main.py
```

#### 打包为独立应用
```bash
cd achievement_report_app
pip install -r requirements.txt
pip install pyinstaller
python build_app.py
```

### 方式二：使用命令行脚本

```bash
# 单文件处理
python process_achievement_data.py

# 批量处理
python process_achievement_data.py --batch
```

## 安装依赖

```bash
# 桌面应用
pip install customtkinter pandas openpyxl

# 命令行脚本
pip install pandas openpyxl
```

## 成绩单格式要求

成绩单Excel需包含：

| 必需信息 | 说明 |
|---------|------|
| 学号 | 学生学号列（长度≥5，数字占比≥80%） |
| 姓名 | 学生姓名列 |
| 平时成绩 | 或"平时"、"平时分" |
| 期末成绩 | 或"期末"、"期末考试" |
| 总成绩 | 或"总评成绩"、"总分"、"成绩"、"总评" |

### 班级识别（可选，自动识别优先级）

1. 单元格内容：`行政班：XXX` 或 `班级：XXX`
2. 工作表名称：如 `9007851-0001_音乐2212` → `音乐2212`
3. 文件名：如 `软件工程2301_成绩.xlsx` → `软件工程2301`
4. 列数据：列头为"班级"或"行政班"

特殊状态（缺考、缓考、作弊等）会自动识别并标记。

## 配置参数

| 参数 | 默认值 | 说明 |
|------|--------|------|
| 目标一占比 | 50% | 三项之和须为100% |
| 目标二占比 | 30% | |
| 目标三占比 | 20% | |
| 平时成绩占比 | 30% | 两项之和须为100% |
| 期末成绩占比 | 70% | |
| 达成度期望值 | 0.6 | 范围0-1 |

## 输出报告

生成的报告包含两个工作表：

1. **课程目标达成度计算**
   - 学生成绩和达成率数据
   - 目标1/2/3达成度折线图
   - 总达成度折线图

2. **达成度统计**
   - 达成度分级统计表（>0.8, 0.6-0.8, 0.5-0.6, 0.4-0.5, <0.4）
   - 各目标达成度分布柱状图

## 项目结构

```
├── process_achievement_data.py              # 命令行脚本
├── achievement_report_app/                  # 桌面应用
│   ├── main.py                             # 主程序
│   ├── core/
│   │   ├── config.py                       # 配置类
│   │   └── processor.py                    # 处理逻辑
│   ├── build_app.py                        # 打包脚本
│   └── requirements.txt                    # 依赖
├── CHANGELOG.md                             # 更新日志
├── CLAUDE.md                                # 开发文档
└── 模板_案例_说明_依赖模板的脚本/             # 历史备份
```

## 更新日志

详见 [CHANGELOG.md](./CHANGELOG.md)

## 许可证

MIT License

## 作者

刘祉祁 - 集美大学

---

© 2026 达成度报告生成器 v1.7
