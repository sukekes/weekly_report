自动获取周报数据脚本
https://github.com/mufeng12138/weekly_report

前提
chrome driver下载地址
http://chromedriver.storage.googleapis.com/index.html

配置环境：通用方式（输出所需安装的python包的名称及版本到txt文件中）
分享环境：pip freeze > requirements.txt
载入环境：pip install -r /path/requirements.txt
https://blog.csdn.net/clksjx/article/details/85331173

用途
通过分析excel数据，自动获取周报所需部分数据

用法
1 配置config.xlsx文件，包含文件路径、提交/验证者姓名、考察时间范围等信息
2 在平台上手动下载所需项目的缺陷数据（zip格式文件）
3 运行脚本自动解压，并获取数据
4 可通过调节代码中的status_date来控制提取周报（默认周五统计数据）或日报，也可在配置文件中修改

进阶用法
开启自动化模式，主动获取缺陷数据（excel文件），因平台不稳定，尚未完成

+++++++++++++++我是分割线+++++++++++++++
execute_case脚本（尚未完成）
前提：
1 本人名下有待办用例

步骤
1、测试负责人-创建项目
2、测试负责人-创建测试模块
3、产品经理-新建并发布需求
4、研发负责人-新建计划，关联需求，提测，选择已自检（开发提测模块）