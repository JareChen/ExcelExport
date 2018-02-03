# ExcelExport
基于 .net core 的可扩展 excel 导出工具，现支持导出为制表符格式的 txt 和对应的 cs 脚本。 

### 自定义导出
继承 OutputFormat 类，实现 Generate 方法，Excel 读取的方式可以参考 GameFrameworkOutputFormat 类。
替换 Main 函数中创建的 OutputFormat 对象即可。

### 配置
可配置导出和导出目录，在根目录中的 config.txt 中修改对应的导入导出路径，默认搜索目录下的所有 .xlsx 文件，导出所有的 sheet.
