# 代码说明
1.第三方库的依赖在requirements.txt中 进行 pip install -r requirements.txt即可将环境安装到本地python
2.入口文件是openLabDataTool
# 使用说明



## 一.统计作业

使用场景：根据课程布置的所有作业计算出学员的作业提交率。

首先将课程的每次作业的表格从小鹅通导出，程序将通过这些表格计算出每位同学的作业提交率，最后可以将计算结果导出到指定的表格或新的表格中。



## 二.提取考试成绩

**注意**：从小鹅通导出成绩单时请务必勾选**姓名**字段，默认字段只有昵称没有姓名。

使用场景：将成绩单中的学员成绩提取，可导入到指定的表格，程序会为其新增字段或在已有的“得分”字段下导入成绩，也可以导出生成新的文件。

首先将考试成绩单从小鹅通导出，根据程序提示进行操作即可。



## 三.移除学员

使用场景：批量删除学员，需要指定一个接受删除的表格，还需要指定一个待删除学员的列表，或者待保留学员的列表。

首先指定待处理表格，然后指定参考表格（可选多个文件），程序将会待处理表格中的学员进行移除操作，可以移除参考表格中存在的学员，也可以保留参考表格中存在的学员，移除待处理表比参考表多出的学员。



## 四.将图片转为excel表格

在某些场景中，老师通常只有实物表格，做一些统计工作时不太方便。可以通过给实物表格拍照，将照片导入本程序，即可生成相应的excel文件。但是受限于图像清晰度、识别率、代码逻辑等限制，表格一定程度上会出现识别错误，需要进行人工校对。另外文字识别使用的是腾讯云的sdk，需要联网使用，每月只有1000次的免费使用次数，超过后无法使用。
