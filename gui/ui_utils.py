"""UI工具模块

包含创建表格、加载帮助文本、应用磨砂玻璃效果等UI相关的工具函数
"""

from PyQt5.QtWidgets import QTableWidget, QHeaderView
from PyQt5.QtGui import QFont, QColor, QPalette, QBrush
from PyQt5.QtCore import Qt


def create_table() -> QTableWidget:
    """创建表格控件
    
    Returns:
        QTableWidget对象
    """
    table = QTableWidget()
    table.setShowGrid(True)
    table.setGridStyle(Qt.SolidLine)
    table.setAlternatingRowColors(True)
    table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
    table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
    return table


def load_help_text(help_text_widget):
    """加载使用说明文本
    
    Args:
        help_text_widget: QTextEdit对象
    """
    help_text = """
# 年段一分三率计算系统 - 使用说明

## 一、系统功能

1. 导入Excel成绩文件
2. 计算学生总分、排名和进退步
3. 统计班级一分三率（平均分、及格率、优秀率）
4. 生成单科班级分析和排名
5. 导入教师信息，支持自动识别班级格式
6. 导出格式化的Excel结果报表，包含染色功能

## 二、使用步骤

### 1. 系统启动
- 双击main.py文件或在命令行中运行"python main.py"
- 等待系统加载完成，进入主界面

### 2. 导入成绩
- 点击"导入成绩"按钮
- 选择包含成绩数据的Excel文件
- 等待系统加载数据完成

### 3. 科目设置
- 在"科目设置"标签页中，点击"从数据识别科目"按钮自动识别科目
- 手动调整科目名称、满分、及格线和优秀线
- 点击"应用参数"按钮保存设置

### 4. 教师设置
- 在"教师设置"标签页中，点击"导入教师"按钮导入教师信息
- 或手动在表格中输入教师信息
- 点击"保存教师配置"按钮保存设置

### 5. 执行计算
- 点击"应用参数并计算"按钮
- 等待计算完成，系统会自动更新各标签页的数据

### 6. 查看结果
- 在"学生排名"标签页查看学生的总分、排名和进退步
- 在"班级汇总"标签页查看班级的一分三率统计
- 在"单科分析"标签页查看各学科的班级分析
- 在"单科排名"标签页查看各学科的班级排名

### 7. 导出报表
- 点击"导出Excel"按钮
- 选择保存路径和文件名
- 等待导出完成，系统会提示导出成功

### 8. 其他功能
- "清除内存数据"按钮：清除所有已加载的数据，重置系统状态
- "班级详情"下拉框：选择特定班级查看详细成绩
- "排名科目"下拉框：选择特定科目查看排名

## 三、注意事项

1. 成绩文件必须包含"班级"和"姓名"列
2. 各科目列名应与科目设置中的名称一致
3. 教师信息文件应包含"班级"、"班主任"和各科目教师列
4. 导出的Excel文件可能需要启用编辑才能查看完整格式

## 四、常见问题

1. **导入失败**：检查文件格式是否正确，是否包含必要的列
2. **计算错误**：检查科目设置是否正确，是否有缺失数据
3. **导出失败**：检查文件路径是否存在，是否有写入权限

## 五、版本信息

**版本 0.6**
- 更新配色方案，优化输出表格的染色机制
- 为班级汇总表分数段添加绿色系染色方案
- 优化学生排名表的染色颜色，使层次更加分明
- 为年级排名、班级排名、语数英总分、语数英排名、上次排名、进退步添加不同颜色标识
- 为输出表格添加冻结窗格功能，提高数据可读性
- 整合所有功能，添加染色功能和教师信息自动识别
- 优化染色颜色方案，使分数分布更加直观
- 调整学生排名表的列顺序，将座号放在姓名之前
- 在单科班级排名的总分一栏添加班主任信息
- 增加导入教师过程中的兼容性，自动识别班级中的数字并替换为已分班级的文字形式

最后更新：2026年
    """
    help_text_widget.setPlainText(help_text)


def apply_frosted_glass_effect(window):
    """应用磨砂玻璃效果
    
    Args:
        window: QMainWindow对象
    """
    # 设置全局样式
    window.setStyleSheet("""
        QMainWindow {
            background-color: rgba(240, 240, 240, 200);
        }
        QWidget {
            background-color: rgba(255, 255, 255, 180);
            border-radius: 5px;
        }
        QFrame {
            background-color: rgba(255, 255, 255, 160);
            border: 1px solid rgba(200, 200, 200, 100);
            border-radius: 5px;
        }
        QPushButton {
            background-color: rgba(220, 220, 220, 180);
            border: 1px solid rgba(180, 180, 180, 150);
            border-radius: 3px;
            padding: 5px 10px;
        }
        QPushButton:hover {
            background-color: rgba(200, 200, 200, 200);
        }
        QPushButton:pressed {
            background-color: rgba(180, 180, 180, 200);
        }
        QTableWidget {
            background-color: rgba(255, 255, 255, 180);
            border: 1px solid rgba(200, 200, 200, 100);
            border-radius: 3px;
        }
        QHeaderView::section {
            background-color: rgba(220, 220, 220, 180);
            border: 1px solid rgba(180, 180, 180, 150);
            padding: 5px;
        }
        QTabWidget::pane {
            background-color: rgba(255, 255, 255, 180);
            border: 1px solid rgba(200, 200, 200, 100);
            border-radius: 5px;
        }
        QTabBar::tab {
            background-color: rgba(220, 220, 220, 180);
            border: 1px solid rgba(180, 180, 180, 150);
            border-radius: 3px;
            padding: 5px 10px;
        }
        QTabBar::tab:selected {
            background-color: rgba(255, 255, 255, 200);
        }
    """)
