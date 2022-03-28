# QT-Automatic-security-inspection
自动生成巡检报告
=========
 ![](https://github.com/yuchen714/QT-Automatic-security-inspection/blob/main/images/%E4%BD%BF%E7%94%A8.gif)
 =========
 

文档中提到的“附件-各应用弱密码常见修复方式示例.docx”：
 

使用方式
=====================
90%情况下：

1、下载安全巡检所需的表格:
 

2、框选这些表格，右键打包成.tar或.zip格式

 ![]( https://github.com/yuchen714/QT-Automatic-security-inspection/blob/main/images/%E5%8E%8B%E7%BC%A9.gif)


++++++++++++++++++++++


3、访问：http://192.168.192.77:8080/ 上传压缩包，输入信息，点执行按钮
 ![](https://github.com/yuchen714/QT-Automatic-security-inspection/blob/main/images/%E9%A1%B5%E9%9D%A2.png)

++++++++++++++++++++++


4、等待一分钟左右，然后访问页面上展示的下载连接
  ![](https://github.com/yuchen714/QT-Automatic-security-inspection/blob/main/images/%E4%B8%8B%E8%BD%BD.gif)


++++++++++++++++++++++


5、手动更新下文档目录


把该报告、弱密码详情列表、附件-各应用弱密码常见修复方式示例.docx、三个文档放一起

输出样例位于：
/xunjian-jiaoben/docx/巡检结果样例.docx




高级使用
==
有自定义需求的场景下：
通过以下两个模板文件，可以修改巡检使用的漏洞优先级、模板样式：
  /docx/必修漏洞列表.csv ；
  /docx/template2.docx
 
 
比如增删必修漏洞列表；或增加或者删改报表模板中的某一段。然后在打tar包的时候，将这两个文件也一并打成tar包：
  ![](https://github.com/yuchen714/QT-Automatic-security-inspection/blob/main/images/%E5%8E%8B%E7%BC%A92.png)
 ++++++++++++++++++++++

 
 
 
接下来的操作和前文完全一样，生成的报告就会使用自定义的漏洞优先级、模板样式。


必修漏洞列表.csv的最上面一行优先级最高，最下面一行最低，目前的逻辑是：如果表中的漏洞全都匹配完，匹配上的洞不够10个，就在除此以外的的漏洞中选择最危急或高危，可远程利用、存在exp的漏洞
实现方式

简介
==
代码构成：100% python

后端是基于 python-docx 和 pandas 这两个库的文档处理
   ![](https://github.com/yuchen714/QT-Automatic-security-inspection/blob/main/images/%E4%BB%A3%E7%A0%811.png)

++++++++++++++++++++++


前端使用 bottle ，监听根路径并返回一个html页面，点击页面的按钮后会调用后端api，生成巡检报告
   ![](https://github.com/yuchen714/QT-Automatic-security-inspection/blob/main/images/%E4%BB%A3%E7%A0%812.png)
++++++++++++++++++++++



结果下载使用的是SimpleHTTPServer ，通过随机生成的文件保存路径（时间戳+随机16位字符串），确保每个使用者只能下载自己的报告

一些 to do：
to do-1：+ 调整服务器配置：目前使用的的服务器生成每份报告大概20s，性能消耗待观查
to do-2：+ 自动识别少上传了哪些表格
to do-3：漂亮的内置模板
to do-4：正在持续优化运行日志
