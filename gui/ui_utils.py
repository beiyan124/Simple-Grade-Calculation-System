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
【系统功能】

年段一分三率计算系统用于快速统计班级和年级的考试成绩，自动计算平均分、及格率、优秀率，并生成学生排名和班级评比结果。包含单科班级分析和单科班级排名两个详细页面。

【操作步骤】

1. 导入数据：点击工具栏"导入文件"按钮，选择包含成绩的Excel文件（可多选）。系统会自动识别班级、姓名、座号及各科成绩列。

2. 设置参数：在左侧面板中，为每个科目设置满分、及格线、优秀线（系统会根据科目名自动填充常用默认值，如语数英150/90/120，体育40/0/0等）；可选择是否计算总分并设置各科权重；设置排名规则及缺考处理方式。

3. （可选）进退步计算：若需计算学生进退步，请导入历史总表（需包含姓名和上次排名），并勾选"启用进退步计算"。

4. 配置教师：在"教师配置"选项卡中，可以设置各班级的班主任和各科任课教师，也可导入已有的教师配置表。

5. 执行计算：点击"应用参数并计算"按钮，系统将根据设置计算：
   - 学生排名表（包含座号、语数英总分及排名、上次排名、进退步）
   - 班级汇总表（含总分分数段统计）
   - 单科班级分析（含分数段分布）
   - 单科班级排名（含比差和名次）

6. 查看结果：在对应选项卡中选择科目即可查看详细数据。

7. 导出结果：点击"导出结果"按钮，选择保存位置，系统将生成包含多个工作表的Excel文件，并自动格式化及配色。

【教师配置说明】

- 教师配置表格中，第一列"班级"自动从数据中获取（不可编辑），第二列"班主任"可编辑，后续各科列可输入任课教师姓名。
- 可通过"导入教师配置"按钮，选择Excel文件快速填充。文件格式：第一列为班级名，第二列为班主任，后续各列按科目顺序对应任课教师。
- 修改教师信息后，需要重新计算才能更新到单科班级排名表和班级汇总表。

【注意事项】

- 导入的Excel文件应包含表头，建议列名包含"班级"、"姓名"、"座号"及科目名称（如语文、数学等）。
- 系统会自动识别成绩列（能转换为数字的列），若科目列包含非数值（如缺考标记），可在"缺考处理"中选择是否排除。
- 权重输入格式：用英文逗号分隔，例如"1,1.2,0.8"，顺序与科目顺序一致（按表格从左到右）。
- 同分处理说明：
  * min（占用名次）：同分者占用相同名次，后续名次跳过（如两人并列第1，下一名为第3）。
  * dense（不占用）：同分者占用相同名次，后续名次不跳过（如两人并列第1，下一名为第2）。
  * average（平均名次）：同分者取平均名次。
- 导出Excel文件将自动调整列宽，百分比列显示为百分比格式，并应用配色方案。

【版本信息】

版本 4.8
新增：语数英总分及排名、进退步计算
教师配置功能完整实现
最后更新：2025年
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
