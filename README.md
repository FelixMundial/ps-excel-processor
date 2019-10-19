报表导入导出相关业务需求Excel工具类，提供报表数据拆分与数据验证等功能（原用作PeopleSoft中Appication Engine程序工具包）

> ##### 依赖

基于Java 7+/Apache POI 3.8

> ##### Code Coverage


> ##### 待实现功能

- [ ] 报表单元格样式的动态读取与动态写入功能有缺陷（边框与字体字号数据能够正确读入和写入，但颜色数据无法正确写入）
- [ ] 适配Apache 4.0+ API
- [ ] 当数据量大时，如何优化JVM以处理OOM问题
