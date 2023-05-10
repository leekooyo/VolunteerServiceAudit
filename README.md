# Volunteer Service Audit #
Vsa，一个基于Pyqt5实现的志愿汇审核软件，为最大化降低审核志愿汇过程中出现的不必要的人力。

Vsa 是一款免费开源软件，地址：[https://github.com/leekoyo/VolunteerServiceAudit](https://github.com/leekoyo/VolunteerServiceAudit)

这里是一张GUI预览图：![](https://github.com/leekooyo/VolunteerServiceAudit/blob/main/doc/README.assets/UI.png)

## 源文件要求 Requirements ##
- 参与名单：表格属性至少包括姓名、班级，表格第一行为属性名。
- 志愿时长名单：表格属性至少包括姓名、信用时数，信用时数数据类型为数值，表格第一行为属性名。
- 无论是哪一个表格，请在使用前确保已经合并了重复人员。如，在一个活动中张三在上午和下午都签了志愿汇，但是哪一个时长都没有超过时长阈值，而合并后的阈值超过了时长阈值，此时会出现漏网之鱼；在1班和2班中都有一个叫张三的人，他们都签了同一个志愿汇活动，此时你应该<del>暴打活动的组织者</del>人工审核这两个志愿汇。
- 无论是哪一个表格，请确保你已经清除了单元格格式，否则会出现异次元的未知错误。

## 源文件示例 Demo ##
详见Demo中的示例文件。

## 安装 Installation ##
[https://github.com/leekooyo/VolunteerServiceAudit/releases](https://github.com/leekooyo/VolunteerServiceAudit/releases)

## 使用 Start ##
首次使用请先打开installer.bat，之后只需要双击Vsa.bat即可运行。

## 联系我们 Contact ##
QQ：815108660
