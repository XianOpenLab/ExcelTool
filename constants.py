MENU_MAIN = """
************************************
按1，统计"作业率"及"得分"
按2，批量删除学员
按3，表格照片转为excel文件
按0，退出程序
***********************************"""

ENTER_CONTINUE = "\r\n*准备就绪后,按enter键继续..."
MENU_MODIFY = f"""
****************************************
*"小鹅通"导出的作业放在->作业<-文件夹
*"小鹅通"导出的成绩单放在->考试<-文件夹
*添加"作业率"和"得分"的excel文件
    1.数据将添加到"结课表"的"作业率""得分"字段，
    保证表名、字段名存在
    2.该文件必须与本程序在同一文件夹
*确保"作业"和"考试"文件夹没有上次使用的数据残留
{ENTER_CONTINUE}
****************************************
"""
MENU_OCR = f"""
****************************************
*将需要转换的图片放在->图片<-文件夹中
{ENTER_CONTINUE}
****************************************
"""
MENU_DELETE = f"""
*************************************
*将要删除的学员的excel放在"删除"文件夹
*excel中至少要有"姓名"列
{ENTER_CONTINUE}
*************************************"""

