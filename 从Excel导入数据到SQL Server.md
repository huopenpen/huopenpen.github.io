## 从Excel导入数据到SQL Server

[TOC]

### SQL Server 2012数据库客户端

1. ![ggp_20171019_161739](http://md.lanpenzi.cn//2020/04/27/ggp_20171019_161739.png-lanpenzi)
2. ![ggp_20171019_161422](http://md.lanpenzi.cn//2020/04/27/ggp_20171019_161422.png-lanpenzi)
3. ![ggp_20171019_161435](http://md.lanpenzi.cn//2020/04/27/ggp_20171019_161435.png-lanpenzi)
4. ![ggp_20171019_161500](http://md.lanpenzi.cn//2020/04/27/ggp_20171019_161500.png-lanpenzi)
5. ![ggp_20171019_161545](http://md.lanpenzi.cn//2020/04/27/ggp_20171019_161545.png-lanpenzi)
6. ![ggp_20171019_161641](http://md.lanpenzi.cn//2020/04/27/ggp_20171019_161641.png-lanpenzi)
7. ![ggp_20171019_161651](http://md.lanpenzi.cn//2020/04/27/ggp_20171019_161651.png-lanpenzi)
8. ![ggp_20171019_161705](http://md.lanpenzi.cn//2020/04/27/ggp_20171019_161705.png-lanpenzi)
9. 注意：出现未在本地计算机上注册“Microsoft.ACE.OLEDB.12.0”提供程序的错误，安装[控件](http://download.microsoft.com/download/7/0/3/703ffbcb-dc0c-4e19-b0da-1463960fdcdb/AccessDatabaseEngine.exe)。



### SQL Server 2005数据库客户端

下面是从**Microsoft Office Excel 2007**电子数据表导入数据到**SQL Server 2005**数据库的逐步指导。

***==下面是关键步骤==***

1. WIN+R（DTSWIZARD）
2. Office 12.0 OLE DB（属性）
3. all - Data Source（文件路径）- Extended Properties（Excel 12.0）

示例表为：`C:\Users\W2\Desktop\数据字典\李国栋\山西版ICD10.xlsx`

***==具体操作==***　　

点击**开始**并选择运行并输入**CMD**然后在命令提示符里输入**DTSWIZARD**。

SQL Server 导入和导出向导的欢迎界面将显示出来，如下图所示：

![AEF641E5-5C69-41A6-8BD2-6F40CFF4FF2E](http://md.lanpenzi.cn//2020/04/27/AEF641E5-5C69-41A6-8BD2-6F40CFF4FF2E.jpg-lanpenzi)

点击下一步按钮，它将进入选择数据源向导界面。

用户应该选择数据源为Microsoft Office 12.0 Access Database Engine OLE DB Provider 然后在向导界面中点击属性…按钮，它将弹出数据链接属性界面。

在所有标签页中，双击数据源属性值并输入电子数据表的位置，例如`C:\Users\W2\Desktop\数据字典\李国栋\山西版ICD10.xlsx`作为导入数据的数据源的Microsoft Office Excel 2007文件名称和路径。

然后双击扩展属性并选择Excel 12.0作为属性值。

![CCDFE4F3-E42C-4CAD-B1DB-B32C538D201A](http://md.lanpenzi.cn//2020/04/27/CCDFE4F3-E42C-4CAD-B1DB-B32C538D201A.jpg-lanpenzi)

![CF8E9F22-588D-4B5B-A64C-8FE4AB61E8F6](http://md.lanpenzi.cn//2020/04/27/CF8E9F22-588D-4B5B-A64C-8FE4AB61E8F6.jpg-lanpenzi)

到Microsoft Office Excel 2007的连接可以通过点击测试连接按钮来进行测试，如下图所示：

![A67FBDBB-1386-4908-9D49-417C22C98267](http://md.lanpenzi.cn//2020/04/27/A67FBDBB-1386-4908-9D49-417C22C98267.jpg-lanpenzi)

在下一个页面中，数据源需要选为SQL Native Client，因为数据将导入到SQL Server 2005。

然后你需要选择数据所要导入的服务器名称，并需要配置合适的验证模式，它之后跟着数据库名称。

![476793F2-FD0C-4610-94E0-2D1173E5228C](http://md.lanpenzi.cn//2020/04/27/476793F2-FD0C-4610-94E0-2D1173E5228C.jpg-lanpenzi)

　　在这个例子中，我们将使用windows验证连接到本地SQL Server实例，所使用的数据库将是ImportExcel。

在Specify Table Copy or Query(指定表复制或查询)向导界面中，选择copy data from one or more tables or views(从一个或多个表或视图复制数据)选项，并继续这个向导到下一个界面。

![3FD246CF-B6F1-40A5-B858-9283E339FF17](http://md.lanpenzi.cn//2020/04/27/3FD246CF-B6F1-40A5-B858-9283E339FF17.jpg-lanpenzi)

在Select Source Table and Views(选择源表和视图)向导界面中，用户需要在源中选择雇员电子数据表，然后在目标中就可以看到ImportExcel.dbo.Employee了。

之后点击Edit Mappings…(编辑匹配…)，扫描电子数据表中的可用数据，如果数据类型与SQL Server所建议的不同的话那么指定数据类型。

![D3EF7E44-46EC-4A1C-852A-0C8F25E09902](http://md.lanpenzi.cn//2020/04/27/D3EF7E44-46EC-4A1C-852A-0C8F25E09902.jpg-lanpenzi)

　　在Save and Execute Package(保存和执行包)向导界面中，有两个选项叫做Execute Immediately(立即执行)和Save SSIS Package as file system(保存SSIS包为文件系统)。你可以选择任何一个选项然后点击Finish(完成)按钮来运行和结束这个包配置。

