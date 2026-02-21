"""
年段一分三率计算系统 - 主程序入口

此文件作为系统的启动点，负责初始化PyQt应用程序并启动GradeAnalyzerApp主窗口。
系统功能包括：
1. 导入Excel成绩文件
2. 计算学生总分、排名和进退步
3. 统计班级一分三率（平均分、及格率、优秀率）
4. 生成单科班级分析和排名
5. 导出格式化的Excel结果报表

作者：Beiyan/deepseek/Trae
版本：4.8
"""

# 导入必要的模块
import sys  # 导入sys模块，用于处理命令行参数和程序退出
from PyQt5.QtWidgets import QApplication, QStyleFactory
from gui import GradeAnalyzerApp  # 导入PyQt版本的GradeAnalyzerApp类


def main():
    """
    初始化应用程序并启动主窗口
    
    此函数执行以下操作：
    1. 创建QApplication应用程序实例
    2. 设置应用程序样式为Fusion（现代风格）
    3. 创建GradeAnalyzerApp主窗口实例
    4. 显示主窗口
    5. 进入应用程序主事件循环，等待用户交互
    """
    try:
        # 创建QApplication应用程序实例
        # sys.argv包含命令行参数列表
        app = QApplication(sys.argv)
        
        # 设置应用程序样式为Fusion，提供现代、跨平台的外观
        app.setStyle(QStyleFactory.create('Fusion'))
        
        # 创建GradeAnalyzerApp主窗口实例
        # PyQt版本的应用程序不需要传递根窗口参数
        window = GradeAnalyzerApp()
        
        # 显示主窗口
        window.show()
        
        # 进入应用程序主事件循环
        # 此函数会阻塞，直到用户关闭窗口
        # sys.exit()确保程序以正确的退出代码结束
        sys.exit(app.exec_())
    except ImportError as e:
        print(f"导入错误: {str(e)}")
        print("请确保已安装所有必要的依赖包，如PyQt5、pandas、openpyxl等")
        input("按回车键退出...")
    except Exception as e:
        print(f"应用程序启动失败: {str(e)}")
        input("按回车键退出...")


if __name__ == "__main__":
    """
    程序入口点检查
    
    当此文件被直接执行时（而不是被导入为模块），
    条件为真，执行main()函数启动应用程序。
    """
    # 调用main函数，启动年段一分三率计算系统
    main()