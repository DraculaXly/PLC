### 一直想研究下WinCC的报表展示，最近正好有时间，就花了两天的时间跟着论坛学习起来啦~
### [西门子论坛地址](http://www.ad.siemens.com.cn/club/bbs/post.aspx?a_id=1364128&b_id=5&s_id=0&num=36&type=elite#)
### 下面是学习成果，里面一小部分代码有修改，并且实际执行遇到的问题也列了出来。有了这两个基础，只要修改其中的部分代码即可实现定制化报表需求。
- [WinCC & Excel](https://github.com/DraculaXly/PLC/tree/master/WinCC%20%26%20Excel)
  - 这个主要是用将过程归档值通过VB脚本读取出来，写入到Excel中。
  - 实际执行过程中可能会遇到DateTimePicker控件无法使用，可以通过下载VB Studio安装后解决。
- [WinCC & ListView](https://github.com/DraculaXly/PLC/tree/master/WinCC%20%26%20ListView)
  - 这个主要是用WinCC的Active X控件，将过程归档值通过VB脚本读取出来，展示在ListView上。
  - 这个实际执行过程时，ListView的View属性需要设置为3，当然了，也可以通过脚本赋值。
- [WinCC & UserArchieve](https://github.com/DraculaXly/PLC/tree/master/WinCC%20%26%20UserArchieve)
  - 每天自动导出UA里面的数据并清空，并自动恢复到之前页面
